import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  ShieldAlert,
  Download,
  Upload,
  RotateCcw,
  CheckCircle2,
  AlertTriangle,
  Sparkles,
  Bot,
  Save,
  Search,
  X,
  Plus,
  Trash2,
  Table,
  FileSpreadsheet,
  ChevronLeft,
  ChevronRight,
  Filter,
  Layers,
  HelpCircle,
  Eye,
  Check,
  RefreshCw
} from "lucide-react";
import * as XLSX from "xlsx";
import { GoogleGenAI } from "@google/genai";
import { parseRobustNumber, parse2DArrayWithSmartHeader } from "../utils/sharedHelpers";
import { EnterpriseAuditAndBctcBridge } from "./EnterpriseAuditAndBctcBridge";

export interface UniversalLogicRule {
  id: string;
  name: string;
  expression: string;
  description: string;
  severity: "error" | "warning";
  enabled: boolean;
  createdAt: string;
}

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
  [key: string]: any;
}

const LOCAL_STORAGE_RULES_KEY = "vtong_gso_expression_rules";

// Hàm chuyển biểu thức GSO sang Javascript an toàn
function parseGsoExpressionToJs(expr: string, availableColumns: string[]) {
  if (!expr || !expr.trim()) return null;

  let js = expr.trim();

  // 1. Thu thập tất cả các token [Cột]
  const colTokens: string[] = [];
  js = js.replace(/\[([^\]]+)\]/g, (_, col) => {
    const idx = colTokens.length;
    colTokens.push(col.trim());
    return `__COL_${idx}__`;
  });

  // 2. Chuyển đổi các toán tử ngữ nghĩa tiếng Việt
  js = js.replace(/(__COL_\d+__)\s+không có giá trị/gi, " _isEmpty($1) ");
  js = js.replace(/(__COL_\d+__)\s+có giá trị/gi, " _isNotEmpty($1) ");
  js = js.replace(/(__COL_\d+__)\s+có chứa\s+('[^']*'|"[^"]*"|__COL_\d+__|\w+)/gi, (_, col, val) => {
    return ` _contains(${col}, ${val}) `;
  });

  // Hỗ trợ thêm toán tử tiếng Anh tương đương
  js = js.replace(/(__COL_\d+__)\s+is\s+null/gi, " _isEmpty($1) ");
  js = js.replace(/(__COL_\d+__)\s+is\s+not\s+null/gi, " _isNotEmpty($1) ");

  // 3. Toán tử logic
  js = js.replace(/\s+và\s+/gi, " && ");
  js = js.replace(/\s+hoặc\s+/gi, " || ");
  js = js.replace(/\s+AND\s+/g, " && ");
  js = js.replace(/\s+OR\s+/g, " || ");

  // 4. Toán tử so sánh: <> thành !==, = thành == (cẩn thận không chạm <=, >=)
  js = js.replace(/<>/g, " !== ");
  js = js.replace(/(?<![<>=!])=(?!=)/g, " == ");

  // 5. Thay thế token __COL_idx__ thành hàm lấy giá trị an toàn
  colTokens.forEach((colName, idx) => {
    const regex = new RegExp(`__COL_${idx}__`, "g");
    const escaped = colName.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    js = js.replace(regex, `_val(row, '${escaped}')`);
  });

  return { js, usedColumns: colTokens };
}

// Helpers an toàn cho hàm eval
const EVAL_HELPERS = {
  _isEmpty: (v: any) => v === null || v === undefined || String(v).trim() === "",
  _isNotEmpty: (v: any) => v !== null && v !== undefined && String(v).trim() !== "",
  _contains: (a: any, b: any) =>
    String(a ?? "").toLowerCase().includes(String(b ?? "").toLowerCase()),
  _val: (row: any, col: string) => {
    if (!row) return "";
    let v = row[col];
    if (v === undefined || v === null) {
      const key = Object.keys(row).find(k => k.toLowerCase() === col.toLowerCase());
      if (key) v = row[key];
    }
    if (v === null || v === undefined || v === "") return "";
    const num = Number(String(v).replace(/,/g, ""));
    return !isNaN(num) && String(v).trim() !== "" ? num : v;
  }
};

export const LogicChecking: React.FC<LogicCheckingProps> = ({
  mainData = [],
  columns = [],
  setMainData,
  setColumns,
  fileName = "",
  setFileName,
  mapping
}) => {
  // Phân hệ chính: "enterprise_bctc" (Kiểm tra lỗi DN & Nạp BCTC) hoặc "query_builder" (Tra cứu dữ liệu & Biểu thức)
  const [mainModuleTab, setMainModuleTab] = useState<"enterprise_bctc" | "query_builder">("enterprise_bctc");

  // Dữ liệu làm việc (đồng bộ từ trang chính hoặc file nạp riêng)
  const [localData, setLocalData] = useState<any[]>(mainData);
  const [localColumns, setLocalColumns] = useState<string[]>(columns);
  const [localFileName, setLocalFileName] = useState<string>(fileName || "Dữ liệu nạp từ hệ thống");
  const [isUsingCustomFile, setIsUsingCustomFile] = useState(false);

  // Quản lý workbook nạp từ Excel để người dùng chọn Sheet / Phiếu điều tra thực tế
  const [uploadedWb, setUploadedWb] = useState<XLSX.WorkBook | null>(null);
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [currentSheetName, setCurrentSheetName] = useState<string>("");

  const fileInputRef = useRef<HTMLInputElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  // Đồng bộ khi mainData ngoài thay đổi (nếu không đang mở file riêng)
  useEffect(() => {
    if (!isUsingCustomFile) {
      setLocalData(mainData);
      setLocalColumns(columns);
      setLocalFileName(fileName || "Dữ liệu nạp từ hệ thống");
    }
  }, [mainData, columns, fileName, isUsingCustomFile]);

  // Bộ lọc hành chính
  const [selectedTinh, setSelectedTinh] = useState("33. Tỉnh Hưng Yên");
  const [selectedCoSo, setSelectedCoSo] = useState("Tất cả");
  const [selectedXa, setSelectedXa] = useState("Chọn xã");

  // DANH SÁCH CHỈ TIÊU HOÀN TOÀN ĐỘNG DỰA TRÊN CỘT THỰC TẾ CỦA PHIẾU / FILE
  const [searchIndicator, setSearchIndicator] = useState("");
  const [selectedIndicator, setSelectedIndicator] = useState<string>("");

  const allIndicators = useMemo(() => {
    if (!localColumns || localColumns.length === 0) return [];

    return localColumns.map((col, idx) => {
      // Phân tích kiểu dữ liệu từ 50 dòng mẫu
      const samples = localData
        .slice(0, 50)
        .map(r => r[col])
        .filter(v => v !== null && v !== undefined && String(v).trim() !== "");

      const isNumeric = samples.length > 0 && samples.every(v => !isNaN(Number(String(v).replace(/,/g, ""))));

      return {
        code: col,
        label: `${idx + 1}. ${col}`,
        isNumeric,
        count: samples.length
      };
    });
  }, [localColumns, localData]);

  const filteredIndicators = useMemo(() => {
    if (!searchIndicator.trim()) return allIndicators;
    const term = searchIndicator.toLowerCase();
    return allIndicators.filter(
      item => item.code.toLowerCase().includes(term) || item.label.toLowerCase().includes(term)
    );
  }, [allIndicators, searchIndicator]);

  // Đặt mặc định chỉ tiêu đầu tiên khi có dữ liệu mới
  useEffect(() => {
    if (allIndicators.length > 0) {
      if (!selectedIndicator || !localColumns.includes(selectedIndicator)) {
        setSelectedIndicator(allIndicators[0].code);
      }
    } else {
      setSelectedIndicator("");
    }
  }, [allIndicators, selectedIndicator, localColumns]);

  // CÁC CÂU GỢI Ý ĐƯỢC SINH TỰ ĐỘNG THEO TIÊU ĐỀ CỘT CỦA PHIẾU ĐANG NẠP (KHÔNG GÁN CỐ ĐỊNH)
  const dynamicSuggestions = useMemo(() => {
    if (!localColumns || localColumns.length === 0) return [];

    const suggestions: { title: string; prompt: string; expression: string }[] = [];

    // Phân loại cột số và cột chuỗi từ dữ liệu thực tế
    const numericCols: string[] = [];
    const textCols: string[] = [];

    localColumns.forEach(col => {
      const samples = localData
        .slice(0, 40)
        .map(r => r[col])
        .filter(v => v !== null && v !== undefined && String(v).trim() !== "");

      const isNum = samples.length > 0 && samples.every(v => !isNaN(Number(String(v).replace(/,/g, ""))));
      if (isNum) numericCols.push(col);
      else textCols.push(col);
    });

    // 1. Kiểm tra để trống cho cột mã / tên / văn bản
    if (textCols.length > 0) {
      const col = textCols[0];
      suggestions.push({
        title: `${col} để trống`,
        prompt: `${col} không có giá trị`,
        expression: `[${col}] không có giá trị`
      });
    }

    // 2. Kiểm tra giá trị âm cho các cột số
    if (numericCols.length > 0) {
      const col = numericCols[0];
      suggestions.push({
        title: `${col} nhỏ hơn 0`,
        prompt: `${col} nhỏ hơn 0`,
        expression: `[${col}] < 0`
      });

      suggestions.push({
        title: `${col} bằng 0`,
        prompt: `${col} = 0`,
        expression: `[${col}] = 0`
      });
    }

    // 3. So sánh tương quan giữa 2 cột số thực tế trong phiếu
    if (numericCols.length >= 2) {
      const c1 = numericCols[0];
      const c2 = numericCols[1];
      suggestions.push({
        title: `${c1} > ${c2}`,
        prompt: `${c1} lớn hơn ${c2}`,
        expression: `[${c1}] > [${c2}]`
      });
    }

    // 4. Cột chuỗi tiếp theo để trống
    if (textCols.length >= 2) {
      const col = textCols[1];
      suggestions.push({
        title: `${col} để trống`,
        prompt: `${col} không có giá trị`,
        expression: `[${col}] không có giá trị`
      });
    }

    // 5. Cột số thứ hai nhỏ hơn 0
    if (numericCols.length >= 2) {
      const col = numericCols[1];
      suggestions.push({
        title: `${col} nhỏ hơn 0`,
        prompt: `${col} nhỏ hơn 0`,
        expression: `[${col}] < 0`
      });
    }

    return suggestions.slice(0, 6);
  }, [localColumns, localData]);

  // Khoảng giá trị thực tế của chỉ tiêu đang chọn (trích xuất trực tiếp từ data)
  const previewValues = useMemo(() => {
    if (!selectedIndicator || localData.length === 0) {
      return [];
    }

    const setVals = new Set<string>();
    for (const row of localData) {
      const v = row[selectedIndicator];
      if (v !== null && v !== undefined && String(v).trim() !== "") {
        setVals.add(String(v));
        if (setVals.size >= 50) break;
      }
    }

    return Array.from(setVals);
  }, [selectedIndicator, localData]);

  // BIỂU THỨC ĐIỀU KIỆN
  const [expression, setExpression] = useState<string>("");
  const [queryError, setQueryError] = useState<string>("");
  const [isQueried, setIsQueried] = useState<boolean>(false);
  const [usedColumnsInQuery, setUsedColumnsInQuery] = useState<string[]>([]);

  // AI Assistant Bar
  const [aiPrompt, setAiPrompt] = useState<string>("");
  const [isAiGenerating, setIsAiGenerating] = useState<boolean>(false);
  const [aiMessage, setAiMessage] = useState<{ type: "success" | "error"; text: string } | null>(null);

  // Danh sách quy tắc đã lưu (khởi tạo từ localStorage, không gán sẵn quy tắc giả)
  const [savedRules, setSavedRules] = useState<UniversalLogicRule[]>(() => {
    try {
      const saved = localStorage.getItem(LOCAL_STORAGE_RULES_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });

  const [showSavedRulesModal, setShowSavedRulesModal] = useState<boolean>(false);

  // Tự động lưu quy tắc vào localStorage
  useEffect(() => {
    try {
      localStorage.setItem(LOCAL_STORAGE_RULES_KEY, JSON.stringify(savedRules));
    } catch {}
  }, [savedRules]);

  // Chèn chuỗi hoặc toán tử vào textarea tại vị trí con trỏ
  const insertIntoExpression = (textToInsert: string) => {
    const textarea = textareaRef.current;
    if (!textarea) {
      setExpression(prev => (prev ? `${prev} ${textToInsert}` : textToInsert));
      return;
    }

    const startPos = textarea.selectionStart ?? expression.length;
    const endPos = textarea.selectionEnd ?? expression.length;
    const before = expression.substring(0, startPos);
    const after = expression.substring(endPos);

    const prefixSpace = before.length > 0 && !before.endsWith(" ") && !before.endsWith("(") ? " " : "";
    const suffixSpace = after.length > 0 && !after.startsWith(" ") && !after.startsWith(")") ? " " : " ";

    const newExpr = `${before}${prefixSpace}${textToInsert}${suffixSpace}${after}`;
    setExpression(newExpr);

    setTimeout(() => {
      textarea.focus();
      const newCursor = startPos + prefixSpace.length + textToInsert.length + suffixSpace.length;
      textarea.setSelectionRange(newCursor, newCursor);
    }, 10);
  };

  // Nạp lại mặc định (Clear biểu thức & kết quả)
  const handleResetQuery = () => {
    setExpression("");
    setQueryError("");
    setIsQueried(false);
    setUsedColumnsInQuery([]);
    setAiMessage(null);
  };

  // Thực hiện tra cứu / quét dữ liệu theo biểu thức
  const [matchedIndices, setMatchedIndices] = useState<number[]>([]);

  const handleExecuteQuery = () => {
    if (!expression.trim()) {
      setQueryError("Vui lòng nhập hoặc chọn biểu thức điều kiện tra cứu!");
      return;
    }

    const parsed = parseGsoExpressionToJs(expression, localColumns);
    if (!parsed) {
      setQueryError("Biểu thức không hợp lệ!");
      return;
    }

    setQueryError("");
    setUsedColumnsInQuery(parsed.usedColumns);

    try {
      const evalFn = new Function(
        "row",
        "_isEmpty",
        "_isNotEmpty",
        "_contains",
        "_val",
        `try { return Boolean(${parsed.js}); } catch(e) { return false; }`
      );

      const matched: number[] = [];
      localData.forEach((row, idx) => {
        const isMatch = evalFn(
          row,
          EVAL_HELPERS._isEmpty,
          EVAL_HELPERS._isNotEmpty,
          EVAL_HELPERS._contains,
          EVAL_HELPERS._val
        );
        if (isMatch) matched.push(idx);
      });

      setMatchedIndices(matched);
      setIsQueried(true);
      setCurrentPage(1);
    } catch (err: any) {
      setQueryError(`Lỗi cú pháp biểu thức: ${err.message || "Vui lòng kiểm tra lại phép toán hoặc tên cột"}`);
    }
  };

  // Trợ lý AI sinh biểu thức từ khẩu lệnh tiếng Việt
  const handleGenerateAiExpression = async () => {
    if (!aiPrompt.trim()) {
      setAiMessage({ type: "error", text: "Vui lòng nhập câu lệnh hoặc yêu cầu tra cứu bằng tiếng Việt!" });
      return;
    }

    if (localColumns.length === 0) {
      setAiMessage({ type: "error", text: "Chưa có dữ liệu phiếu để đối chiếu tiêu đề cột. Vui lòng nạp file Excel trước!" });
      return;
    }

    setIsAiGenerating(true);
    setAiMessage(null);

    const apiKey =
      (import.meta as any).env?.VITE_GEMINI_API_KEY ||
      localStorage.getItem("VITE_GEMINI_API_KEY") ||
      "AQ.Ab8RN6KLde-z8EMqDKTLwQqUG4GHLtbKg6-Isumal41h6SHiMg";

    try {
      let generatedExpression = "";
      let explanation = "";

      // 1. Gọi qua endpoint máy chủ nội bộ
      try {
        const serverRes = await fetch("/api/ai/generate-expression", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            prompt: aiPrompt,
            columns: localColumns,
            customApiKey: apiKey
          })
        });

        if (serverRes.ok) {
          const serverData = await serverRes.json();
          if (serverData.success && serverData.expression) {
            generatedExpression = serverData.expression;
            explanation = serverData.explanation || "";
          }
        }
      } catch {}

      // 2. Dự phòng: SDK client-side
      if (!generatedExpression && apiKey) {
        const ai = new GoogleGenAI({ apiKey });
        const systemPrompt = `Bạn là chuyên gia tra cứu dữ liệu thống kê GSO. Nhiệm vụ của bạn là đọc yêu cầu tra cứu/kiểm tra bằng tiếng Việt và chuyển thành biểu thức điều kiện chính xác theo các cột thực tế của phiếu:
- Tên cột nằm trong ngoặc vuông: [TênCột]. Danh sách cột thực tế của phiếu: [${localColumns.join(", ")}]
- Các toán tử: =, <>, <, <=, >, >=, có chứa, không có giá trị, có giá trị, và, hoặc, +, -, *, /, (, )
Yêu cầu: "${aiPrompt}"
Chỉ trả về JSON: { "expression": "biểu thức", "explanation": "giải thích" }`;

        const res = await ai.models.generateContent({
          model: "gemini-2.5-flash",
          contents: systemPrompt
        });

        const text = res.text || "";
        const cleaned = text.replace(/```json/gi, "").replace(/```/g, "").trim();
        const parsed = JSON.parse(cleaned);
        if (parsed.expression) {
          generatedExpression = parsed.expression;
          explanation = parsed.explanation || "";
        }
      }

      if (generatedExpression) {
        setExpression(generatedExpression);
        setAiMessage({
          type: "success",
          text: `Đã dịch biểu thức: ${generatedExpression} ${explanation ? `(${explanation})` : ""}`
        });
      } else {
        setAiMessage({ type: "error", text: "AI không thể nhận diện điều kiện phù hợp với các cột trong phiếu hiện tại." });
      }
    } catch (err: any) {
      setAiMessage({ type: "error", text: `Lỗi gọi AI: ${err?.message || "Không thể kết nối dịch vụ AI"}` });
    } finally {
      setIsAiGenerating(false);
    }
  };

  // Lưu biểu thức hiện tại thành quy tắc logic
  const handleSaveCurrentAsRule = () => {
    if (!expression.trim()) {
      alert("Vui lòng nhập biểu thức trước khi lưu!");
      return;
    }

    const ruleName = prompt("Nhập tên cho quy tắc này:", `Kiểm tra: ${expression.slice(0, 30)}...`);
    if (!ruleName) return;

    const newRule: UniversalLogicRule = {
      id: `rule_${Date.now()}`,
      name: ruleName.trim(),
      expression: expression.trim(),
      description: `Quy tắc tra cứu: ${expression.trim()}`,
      severity: "error",
      enabled: true,
      createdAt: new Date().toLocaleDateString("vi-VN")
    };

    setSavedRules(prev => [newRule, ...prev]);
    alert(`Đã lưu thành công quy tắc: "${newRule.name}"!`);
  };

  // Chuyển đổi Sheet khi file Excel có nhiều Phiếu / Trang tính
  const handleSelectSheet = (sheetName: string) => {
    if (!uploadedWb || !uploadedWb.Sheets[sheetName]) return;
    const ws = uploadedWb.Sheets[sheetName];
    const raw2D: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const parsed = parse2DArrayWithSmartHeader(raw2D);
    if (parsed.data.length > 0) {
      setLocalData(parsed.data);
      setLocalColumns(parsed.columns);
      setCurrentSheetName(sheetName);
      handleResetQuery();
    }
  };

  // Nạp lại file Excel riêng
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        setUploadedWb(wb);
        setAvailableSheets(wb.SheetNames);

        const firstSheet = wb.SheetNames[0];
        setCurrentSheetName(firstSheet);

        const ws = wb.Sheets[firstSheet];
        const raw2D: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

        const parsed = parse2DArrayWithSmartHeader(raw2D);
        if (parsed.data.length > 0) {
          setLocalData(parsed.data);
          setLocalColumns(parsed.columns);
          setLocalFileName(file.name);
          setIsUsingCustomFile(true);
          handleResetQuery();
        } else {
          alert("Tệp không có dữ liệu hợp lệ!");
        }
      } catch (err: any) {
        alert("Lỗi khi đọc tệp Excel: " + (err?.message || "Không thể nạp tệp"));
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  // Quay về dữ liệu gốc của trang tổng quan
  const handleResetToMainData = () => {
    setLocalData(mainData);
    setLocalColumns(columns);
    setLocalFileName(fileName || "Dữ liệu nạp từ hệ thống");
    setIsUsingCustomFile(false);
    setUploadedWb(null);
    setAvailableSheets([]);
    setCurrentSheetName("");
    handleResetQuery();
  };

  // Xuất bảng kết quả ra Excel
  const handleExportResultToExcel = () => {
    const rowsToExport = isQueried ? matchedIndices.map(idx => localData[idx]) : localData;
    if (rowsToExport.length === 0) {
      alert("Không có dòng dữ liệu nào để xuất!");
      return;
    }

    const ws = XLSX.utils.json_to_sheet(rowsToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "KetQuaTraCuu");
    XLSX.writeFile(wb, `KetQuaTraCuu_${Date.now()}.xlsx`);
  };

  // Quản lý hiển thị bảng kết quả bên dưới
  const [tableFilterMode, setTableFilterMode] = useState<"matched_only" | "all">("matched_only");
  const [tableSearchTerm, setTableSearchTerm] = useState<string>("");
  const [rowsPerPage, setRowsPerPage] = useState<number>(25);
  const [currentPage, setCurrentPage] = useState<number>(1);

  const displayedRows = useMemo(() => {
    let indices: number[] = [];

    if (isQueried) {
      if (tableFilterMode === "matched_only") {
        indices = [...matchedIndices];
      } else {
        indices = localData.map((_, i) => i);
      }
    } else {
      indices = localData.map((_, i) => i);
    }

    if (tableSearchTerm.trim()) {
      const term = tableSearchTerm.toLowerCase();
      indices = indices.filter(idx => {
        const row = localData[idx];
        return Object.values(row).some(v => String(v).toLowerCase().includes(term));
      });
    }

    return indices;
  }, [isQueried, tableFilterMode, matchedIndices, localData, tableSearchTerm]);

  const totalPages = Math.max(1, Math.ceil(displayedRows.length / rowsPerPage));
  const paginatedIndices = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return displayedRows.slice(start, start + rowsPerPage);
  }, [displayedRows, currentPage, rowsPerPage]);

  return (
    <div className="space-y-4 pb-8 font-sans text-slate-800">
      {/* THANH ĐIỀU HƯỚNG CẤP CAO - PHÂN HỆ KIỂM TRA LỖI DN / BCTC & TRA CỨU */}
      <div className="bg-white border border-slate-300 p-2 shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-2.5">
        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => setMainModuleTab("enterprise_bctc")}
            className={`px-3.5 py-2 text-xs font-bold flex items-center gap-2 border cursor-pointer transition-all ${
              mainModuleTab === "enterprise_bctc"
                ? "bg-[#286e42] text-white border-[#1d4f2f] shadow-xs ring-1 ring-emerald-500/50"
                : "bg-slate-50 text-slate-700 border-slate-300 hover:bg-slate-100"
            }`}
          >
            <ShieldAlert className={`w-4 h-4 ${mainModuleTab === "enterprise_bctc" ? "text-amber-300" : "text-emerald-700"}`} />
            <span>KIỂM TRA LỖI DN & NẠP BCTC (XML/PDF) ĐIỀN MẪU PHIẾU 01/IO-DN</span>
            <span className="bg-amber-400 text-slate-950 text-[10px] font-black px-1.5 py-0.2">MỚI</span>
          </button>

          <button
            type="button"
            onClick={() => setMainModuleTab("query_builder")}
            className={`px-3.5 py-2 text-xs font-bold flex items-center gap-2 border cursor-pointer transition-all ${
              mainModuleTab === "query_builder"
                ? "bg-[#0d6efd] text-white border-[#0b5ed7] shadow-xs ring-1 ring-sky-500/50"
                : "bg-slate-50 text-slate-700 border-slate-300 hover:bg-slate-100"
            }`}
          >
            <Table className={`w-4 h-4 ${mainModuleTab === "query_builder" ? "text-sky-200" : "text-sky-700"}`} />
            <span>TRA CỨU & BIỂU THỨC QUY TẮC ĐỘNG THEO CỘT PHIẾU ({localColumns.length} CỘT)</span>
          </button>
        </div>

        <div className="text-[11px] text-slate-500 flex items-center gap-2">
          <span className="bg-slate-100 px-2 py-1 border border-slate-200 font-mono">
            {localFileName} ({localData.length.toLocaleString("vi-VN")} dòng)
          </span>
        </div>
      </div>

      {/* HIỂN THỊ PHÂN HỆ KIỂM TRA LỖI DOANH NGHIỆP & NẠP BCTC (IO-2026) */}
      {mainModuleTab === "enterprise_bctc" && (
        <EnterpriseAuditAndBctcBridge
          mainData={localData}
          columns={localColumns}
          fileName={localFileName}
        />
      )}

      {/* HIỂN THỊ PHÂN HỆ TRA CỨU & BIỂU THỨC QUY TẮC ĐỘNG */}
      {mainModuleTab === "query_builder" && (
        <div className="bg-white border border-slate-300 shadow-sm space-y-3 pb-8">
      {/* 1. TIÊU ĐỀ CHÍNH - TRA CỨU DỮ LIỆU PHIẾU ĐIỀU TRA */}
      <div className="text-center pt-3 pb-1 border-b border-slate-200">
        <h2 className="text-base sm:text-lg md:text-xl font-bold uppercase tracking-wider text-[#0d6efd] m-0">
          TRA CỨU DỮ LIỆU PHIẾU ĐIỀU TRA
        </h2>
        <div className="flex flex-wrap items-center justify-center gap-3 text-xs text-slate-500 mt-1">
          <span>
            Phiếu hiện tại: <strong>{localFileName}</strong>{" "}
            {currentSheetName && <span className="text-sky-700 font-bold">(Trang: {currentSheetName})</span>} -{" "}
            <strong>{localData.length.toLocaleString("vi-VN")} dòng</strong>,{" "}
            <strong>{localColumns.length} chỉ tiêu/cột</strong>
          </span>
          {isUsingCustomFile && (
            <span className="bg-amber-100 text-amber-900 border border-amber-300 text-[10px] font-bold px-1.5 py-0.2">
              Tệp kiểm tra riêng
            </span>
          )}
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="text-sky-600 hover:text-sky-800 font-bold underline cursor-pointer border-0 bg-transparent flex items-center gap-1"
          >
            <Upload className="w-3 h-3" /> Nạp tệp phiếu Excel
          </button>
          {isUsingCustomFile && (
            <button
              type="button"
              onClick={handleResetToMainData}
              className="text-emerald-700 hover:text-emerald-900 font-bold underline cursor-pointer border-0 bg-transparent"
            >
              Quay lại dữ liệu gốc
            </button>
          )}
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xlsx,.xls,.csv"
            className="hidden"
          />
        </div>
      </div>

      {/* 2. THANH BỘ LỌC THEO PHIẾU & ĐỊA BÀN THỰC TẾ */}
      <div className="px-3 sm:px-5 space-y-2">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-2.5">
          {/* Tỉnh */}
          <div className="relative border border-sky-400 bg-white px-2.5 pt-2 pb-1 focus-within:border-sky-600">
            <label className="absolute -top-2 left-2 bg-white px-1 text-[11px] font-semibold text-sky-700">
              Tỉnh
            </label>
            <select
              value={selectedTinh}
              onChange={e => setSelectedTinh(e.target.value)}
              className="w-full bg-transparent text-xs text-slate-900 font-medium focus:outline-none cursor-pointer border-0"
            >
              <option value="33. Tỉnh Hưng Yên">33. Tỉnh Hưng Yên</option>
              <option value="01. TP Hà Nội">01. TP Hà Nội</option>
              <option value="79. TP Hồ Chí Minh">79. TP Hồ Chí Minh</option>
              <option value="27. Tỉnh Bắc Ninh">27. Tỉnh Bắc Ninh</option>
              <option value="31. TP Hải Phòng">31. TP Hải Phòng</option>
              <option value="30. Tỉnh Hải Dương">30. Tỉnh Hải Dương</option>
              <option value="35. Tỉnh Hà Nam">35. Tỉnh Hà Nam</option>
              <option value="36. Tỉnh Nam Định">36. Tỉnh Nam Định</option>
              <option value="37. Tỉnh Ninh Bình">37. Tỉnh Ninh Bình</option>
              <option value="34. Tỉnh Thái Bình">34. Tỉnh Thái Bình</option>
            </select>
          </div>

          {/* Thống kê cơ sở */}
          <div className="relative border border-sky-400 bg-white px-2.5 pt-2 pb-1 focus-within:border-sky-600">
            <label className="absolute -top-2 left-2 bg-white px-1 text-[11px] font-semibold text-sky-700">
              Thống kê cơ sở
            </label>
            <select
              value={selectedCoSo}
              onChange={e => setSelectedCoSo(e.target.value)}
              className="w-full bg-transparent text-xs text-slate-900 font-medium focus:outline-none cursor-pointer border-0"
            >
              <option value="Tất cả">Tất cả</option>
              <option value="Chi cục Thống kê TP Hưng Yên">Chi cục Thống kê TP Hưng Yên</option>
              <option value="Chi cục TK khu vực Khoái Châu - Kim Động">Chi cục TK khu vực Khoái Châu - Kim Động</option>
              <option value="Chi cục TK khu vực Tiên Lữ - Phù Cừ">Chi cục TK khu vực Tiên Lữ - Phù Cừ</option>
              <option value="Chi cục TK khu vực Mỹ Hào - Văn Lâm">Chi cục TK khu vực Mỹ Hào - Văn Lâm</option>
              <option value="Chi cục TK khu vực Yên Mỹ - Văn Giang">Chi cục TK khu vực Yên Mỹ - Văn Giang</option>
              <option value="Chi cục Thống kê Ân Thi">Chi cục Thống kê Ân Thi</option>
            </select>
          </div>

          {/* Xã */}
          <div className="relative border border-sky-400 bg-white px-2.5 pt-2 pb-1 focus-within:border-sky-600">
            <label className="absolute -top-2 left-2 bg-white px-1 text-[11px] font-semibold text-sky-700">
              Xã
            </label>
            <select
              value={selectedXa}
              onChange={e => setSelectedXa(e.target.value)}
              className="w-full bg-transparent text-xs text-slate-900 font-medium focus:outline-none cursor-pointer border-0"
            >
              <option value="Chọn xã">Chọn xã</option>
              <option value="Phường Hiến Nam">Phường Hiến Nam</option>
              <option value="Phường Lam Sơn">Phường Lam Sơn</option>
              <option value="Phường An Tảo">Phường An Tảo</option>
              <option value="Phường Lê Lợi">Phường Lê Lợi</option>
              <option value="Phường Minh Khai">Phường Minh Khai</option>
              <option value="Phường Quang Trung">Phường Quang Trung</option>
              <option value="Phường Hồng Châu">Phường Hồng Châu</option>
              <option value="Xã Bảo Khê">Xã Bảo Khê</option>
              <option value="Xã Trung Nghĩa">Xã Trung Nghĩa</option>
              <option value="Xã Liên Phương">Xã Liên Phương</option>
              <option value="Xã Hồng Nam">Xã Hồng Nam</option>
            </select>
          </div>
        </div>

        {/* Hàng 2: Phiếu điều tra (Động theo các Sheet trong file đã nạp) */}
        <div className="relative border border-sky-400 bg-white px-2.5 pt-2 pb-1 focus-within:border-sky-600 flex items-center justify-between">
          <label className="absolute -top-2 left-2 bg-white px-1 text-[11px] font-semibold text-sky-700">
            Phiếu điều tra / Trang tính đang chọn
          </label>
          {availableSheets.length > 1 ? (
            <select
              value={currentSheetName}
              onChange={e => handleSelectSheet(e.target.value)}
              className="w-full bg-transparent text-xs text-slate-900 font-bold focus:outline-none cursor-pointer border-0 pr-6"
            >
              {availableSheets.map((sName, sIdx) => (
                <option key={sName} value={sName}>
                  Phiếu {sIdx + 1}: {sName} ({localFileName})
                </option>
              ))}
            </select>
          ) : (
            <div className="w-full text-xs text-slate-900 font-semibold truncate pr-6 py-0.5">
              {localFileName ? `${localFileName} ${currentSheetName ? `[${currentSheetName}]` : ""}` : "Chưa có file phiếu"}
            </div>
          )}
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="text-sky-600 hover:text-sky-800 cursor-pointer border-0 bg-transparent p-0.5"
            title="Đổi tệp phiếu khác"
          >
            <Upload className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* 3. KHU VỰC TRỢ LÝ AI: GÕ LỆNH TỰ NHIÊN DỊCH THÀNH BIỂU THỨC */}
      <div className="mx-3 sm:mx-5 bg-gradient-to-r from-sky-50 via-indigo-50 to-emerald-50 border border-sky-300 p-2.5 space-y-2">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1">
          <span className="text-xs font-bold text-sky-950 flex items-center gap-1.5">
            <Sparkles className="w-4 h-4 text-indigo-600 animate-pulse" />
            Trợ lý AI dịch khẩu lệnh tự nhiên theo các cột của phiếu điều tra
          </span>
          <span className="text-[11px] text-slate-500 font-medium">
            (Có thể vừa dùng AI dịch, vừa gõ lệnh tự do hoặc nhấp chọn các chỉ tiêu bên dưới)
          </span>
        </div>

        <div className="flex gap-2">
          <div className="relative flex-1">
            <input
              type="text"
              value={aiPrompt}
              onChange={e => setAiPrompt(e.target.value)}
              onKeyDown={e => {
                if (e.key === "Enter") handleGenerateAiExpression();
              }}
              placeholder={
                localColumns.length > 0
                  ? `Nhập yêu cầu kiểm tra (ví dụ: "${localColumns[0]} không có giá trị"${
                      localColumns[1] ? `, "${localColumns[1]} nhỏ hơn 0"` : ""
                    })...`
                  : "Vui lòng nạp file Excel trước để AI đối chiếu theo tiêu đề cột..."
              }
              className="w-full bg-white border border-sky-300 text-xs px-3 py-1.5 text-slate-900 focus:outline-none focus:border-indigo-600 shadow-2xs"
            />
            {aiPrompt && (
              <button
                type="button"
                onClick={() => setAiPrompt("")}
                className="absolute right-2 top-2 text-slate-400 hover:text-slate-700 cursor-pointer border-0 bg-transparent text-xs"
              >
                ✕
              </button>
            )}
          </div>
          <button
            type="button"
            onClick={handleGenerateAiExpression}
            disabled={isAiGenerating || localColumns.length === 0}
            className="bg-indigo-700 hover:bg-indigo-800 disabled:bg-slate-300 text-white font-bold text-xs px-3.5 py-1.5 flex items-center gap-1.5 cursor-pointer border-0 shadow-xs shrink-0"
          >
            <Bot className="w-3.5 h-3.5 text-amber-300" />
            <span>{isAiGenerating ? "AI đang dịch..." : "AI Tạo Biểu Thức"}</span>
          </button>
        </div>

        {/* CÁC CÂU GỢI Ý ĐƯỢC SINH ĐỘNG THEO TIÊU ĐỀ CỘT THỰC TẾ TRONG PHIẾU */}
        <div className="flex flex-wrap items-center gap-1.5 pt-0.5">
          <span className="text-[11px] font-bold text-slate-600">Gợi ý theo tiêu đề cột:</span>
          {dynamicSuggestions.length === 0 ? (
            <span className="text-[11px] text-slate-400 italic">
              (Nạp tệp Excel để tự động hiển thị gợi ý theo các cột của phiếu)
            </span>
          ) : (
            dynamicSuggestions.map((sug, qIdx) => (
              <button
                key={qIdx}
                type="button"
                onClick={() => {
                  setAiPrompt(sug.prompt);
                  setExpression(sug.expression);
                }}
                className="bg-white/90 hover:bg-white text-indigo-900 border border-indigo-200 hover:border-indigo-500 text-[11px] px-2 py-0.5 cursor-pointer font-medium shadow-2xs transition-colors"
                title={`Nhấp để áp dụng ngay biểu thức: ${sug.expression}`}
              >
                {sug.title}
              </button>
            ))
          )}
        </div>

        {aiMessage && (
          <p
            className={`text-xs font-bold m-0 p-1.5 ${
              aiMessage.type === "success"
                ? "bg-emerald-100/70 text-emerald-900 border border-emerald-300"
                : "bg-rose-100/70 text-rose-900 border border-rose-300"
            }`}
          >
            {aiMessage.text}
          </p>
        )}
      </div>

      {/* 4. KHỐI TẠO BIỂU THỨC - 3 CỘT (CHỈ TIÊU HOÀN TOÀN TỪ CỘT CỦA PHIẾU) */}
      <div className="px-3 sm:px-5">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-3">
          {/* CỘT 1: DANH SÁCH CÁC CHỈ TIÊU (THEO ĐÚNG CỘT CỦA PHIẾU ĐANG NẠP) */}
          <div className="lg:col-span-6 border border-slate-300 bg-white flex flex-col h-[270px]">
            <div className="bg-slate-50 px-2.5 py-1.5 border-b border-slate-200 flex items-center justify-between">
              <span className="text-xs font-bold text-slate-900">Danh sách các chỉ tiêu trong phiếu</span>
              <span className="text-[11px] text-slate-500 font-medium">({filteredIndicators.length} chỉ tiêu/cột)</span>
            </div>

            {/* Ô tìm nhanh chỉ tiêu */}
            <div className="p-1.5 border-b border-slate-200 bg-slate-50/50">
              <div className="relative">
                <input
                  type="text"
                  value={searchIndicator}
                  onChange={e => setSearchIndicator(e.target.value)}
                  placeholder="🔍 Tìm nhanh chỉ tiêu theo tên cột..."
                  className="w-full bg-white border border-slate-300 text-xs px-2 py-1 text-slate-800 focus:outline-none focus:border-sky-500"
                />
                {searchIndicator && (
                  <button
                    type="button"
                    onClick={() => setSearchIndicator("")}
                    className="absolute right-2 top-1 text-slate-400 hover:text-slate-700 cursor-pointer border-0 bg-transparent text-xs"
                  >
                    ✕
                  </button>
                )}
              </div>
            </div>

            {/* Danh sách cuộn chỉ tiêu */}
            <div className="flex-1 overflow-y-auto divide-y divide-slate-100 text-xs select-none">
              {filteredIndicators.length === 0 ? (
                <div className="p-4 text-center text-slate-400 italic">
                  Chưa có chỉ tiêu nào. Vui lòng nạp file phiếu điều tra để hiển thị danh sách tiêu đề cột.
                </div>
              ) : (
                filteredIndicators.map((item, idx) => {
                  const isSelected = selectedIndicator === item.code;
                  return (
                    <div
                      key={idx}
                      onClick={() => setSelectedIndicator(item.code)}
                      onDoubleClick={() => insertIntoExpression(`[${item.code}]`)}
                      className={`px-2.5 py-1.5 cursor-pointer flex items-center justify-between transition-colors ${
                        isSelected
                          ? "bg-sky-100 text-sky-950 font-bold border-l-4 border-sky-600"
                          : "hover:bg-slate-50 text-slate-700 font-normal"
                      }`}
                      title="Nhấp đúp để chèn vào biểu thức điều kiện"
                    >
                      <div className="truncate pr-2 flex items-center gap-1.5">
                        <span>{item.label}</span>
                        {item.isNumeric ? (
                          <span className="text-[9px] bg-sky-100 text-sky-800 px-1 font-mono font-normal">số</span>
                        ) : (
                          <span className="text-[9px] bg-slate-100 text-slate-600 px-1 font-mono font-normal">chuỗi</span>
                        )}
                      </div>
                      <button
                        type="button"
                        onClick={e => {
                          e.stopPropagation();
                          insertIntoExpression(`[${item.code}]`);
                        }}
                        className="text-[10px] text-sky-700 bg-sky-50 hover:bg-sky-200 border border-sky-300 px-1.5 py-0.2 font-mono shrink-0 cursor-pointer"
                        title="Chèn chỉ tiêu này"
                      >
                        + Chèn
                      </button>
                    </div>
                  );
                })
              )}
            </div>
          </div>

          {/* CỘT 2: PHÉP TOÁN */}
          <div className="lg:col-span-3 border border-slate-300 bg-white flex flex-col h-[270px]">
            <div className="bg-slate-50 px-2.5 py-1.5 border-b border-slate-200">
              <span className="text-xs font-bold text-slate-900">Phép toán</span>
            </div>

            <div className="p-2 flex-1 flex flex-col justify-between gap-1.5 bg-slate-50/30">
              {/* Hàng 1: = , <> , < , <= */}
              <div className="grid grid-cols-4 gap-1.5">
                {["=", "<>", "<", "<="].map(op => (
                  <button
                    key={op}
                    type="button"
                    onClick={() => insertIntoExpression(op)}
                    className="bg-white hover:bg-sky-50 active:bg-sky-100 text-slate-800 font-mono font-bold text-xs py-1.5 border border-slate-300 hover:border-sky-400 cursor-pointer shadow-2xs transition-colors"
                  >
                    {op}
                  </button>
                ))}
              </div>

              {/* Hàng 2: > , >= */}
              <div className="grid grid-cols-2 gap-1.5">
                {[">", ">="].map(op => (
                  <button
                    key={op}
                    type="button"
                    onClick={() => insertIntoExpression(op)}
                    className="bg-white hover:bg-sky-50 active:bg-sky-100 text-slate-800 font-mono font-bold text-xs py-1.5 border border-slate-300 hover:border-sky-400 cursor-pointer shadow-2xs transition-colors"
                  >
                    {op}
                  </button>
                ))}
              </div>

              {/* Hàng 3: có chứa , không có giá trị , có giá trị */}
              <div className="grid grid-cols-3 gap-1.5">
                {[
                  { text: "có chứa", label: "có chứa" },
                  { text: "không có giá trị", label: "không có giá trị" },
                  { text: "có giá trị", label: "có giá trị" }
                ].map(item => (
                  <button
                    key={item.text}
                    type="button"
                    onClick={() => insertIntoExpression(item.text)}
                    className="bg-white hover:bg-sky-50 active:bg-sky-100 text-slate-800 font-bold text-[11px] py-1.5 px-0.5 border border-slate-300 hover:border-sky-400 cursor-pointer shadow-2xs text-center leading-tight transition-colors"
                  >
                    {item.label}
                  </button>
                ))}
              </div>

              {/* Hàng 4: và , hoặc , ( , ) */}
              <div className="grid grid-cols-4 gap-1.5">
                {["và", "hoặc", "(", ")"].map(op => (
                  <button
                    key={op}
                    type="button"
                    onClick={() => insertIntoExpression(op)}
                    className="bg-white hover:bg-sky-50 active:bg-sky-100 text-slate-800 font-bold text-xs py-1.5 border border-slate-300 hover:border-sky-400 cursor-pointer shadow-2xs transition-colors"
                  >
                    {op}
                  </button>
                ))}
              </div>

              {/* Hàng 5: + , - , * , / */}
              <div className="grid grid-cols-4 gap-1.5">
                {["+", "-", "*", "/"].map(op => (
                  <button
                    key={op}
                    type="button"
                    onClick={() => insertIntoExpression(op)}
                    className="bg-white hover:bg-sky-50 active:bg-sky-100 text-slate-800 font-mono font-bold text-xs py-1.5 border border-slate-300 hover:border-sky-400 cursor-pointer shadow-2xs transition-colors"
                  >
                    {op}
                  </button>
                ))}
              </div>
            </div>
          </div>

          {/* CỘT 3: KHOẢNG GIÁ TRỊ (TRÍCH TỪ CỘT ĐANG CHỌN) */}
          <div className="lg:col-span-3 border border-slate-300 bg-white flex flex-col h-[270px]">
            <div className="bg-slate-50 px-2.5 py-1.5 border-b border-slate-200 flex items-center justify-between">
              <span className="text-xs font-bold text-slate-900">Khoảng giá trị</span>
              <span className="text-[10px] text-slate-500 font-mono font-semibold truncate max-w-[120px]">
                {selectedIndicator || "Chọn chỉ tiêu"}
              </span>
            </div>

            <div className="flex-1 overflow-y-auto divide-y divide-slate-100 text-xs p-1">
              {!selectedIndicator ? (
                <div className="p-3 text-center text-slate-400 italic">
                  Vui lòng chọn 1 chỉ tiêu ở danh sách bên trái để xem giá trị.
                </div>
              ) : previewValues.length === 0 ? (
                <div className="p-3 text-center text-slate-400 italic">
                  Cột này không có giá trị hoặc toàn bộ ô đang để trống.
                </div>
              ) : (
                <>
                  <div className="px-1 py-0.5 text-[10px] text-slate-400 uppercase font-semibold">
                    Giá trị thực tế trong cột ({previewValues.length}):
                  </div>
                  {previewValues.map((val, idx) => (
                    <div
                      key={idx}
                      onClick={() => insertIntoExpression(isNaN(Number(val)) ? `'${val}'` : val)}
                      className="px-2 py-1 hover:bg-slate-100 cursor-pointer flex items-center justify-between font-mono text-[11px] text-slate-700"
                      title="Nhấp để chèn giá trị này vào biểu thức"
                    >
                      <span className="truncate pr-1">{val}</span>
                      <span className="text-[10px] text-sky-600 bg-sky-50 px-1 border border-sky-200 shrink-0">
                        + Chèn
                      </span>
                    </div>
                  ))}
                </>
              )}
            </div>
          </div>
        </div>

        {/* 5. KHUNG NHẬP BIỂU THỨC ĐIỀU KIỆN */}
        <div className="mt-3 space-y-1.5">
          <div className="relative border border-sky-400 bg-white focus-within:border-sky-600 focus-within:ring-1 focus-within:ring-sky-400">
            <textarea
              ref={textareaRef}
              rows={3}
              value={expression}
              onChange={e => {
                setExpression(e.target.value);
                setQueryError("");
              }}
              placeholder="Biểu thức điều kiện... (Nhấp đúp chỉ tiêu hoặc toán tử ở trên, gõ lệnh tự do, hoặc dùng AI phía trên)"
              className="w-full p-2.5 text-xs sm:text-sm font-mono text-slate-900 bg-transparent focus:outline-none resize-y"
            />
            {expression && (
              <button
                type="button"
                onClick={() => setExpression("")}
                className="absolute top-2 right-2 text-slate-400 hover:text-rose-600 cursor-pointer border-0 bg-transparent text-xs p-1"
                title="Xóa biểu thức"
              >
                ✕ Xóa
              </button>
            )}
          </div>

          {queryError && (
            <p className="text-xs font-bold text-rose-600 m-0 bg-rose-50 border border-rose-200 p-1.5">
              ⚠️ {queryError}
            </p>
          )}
        </div>

        {/* 6. CÁC NÚT THAO TÁC CHÍNH */}
        <div className="mt-3 flex flex-col sm:flex-row items-center justify-between gap-2.5">
          {/* NÚT THỰC HIỆN LẠI (MÀU ĐỎ NHƯ ẢNH) */}
          <button
            type="button"
            onClick={handleResetQuery}
            className="w-full sm:w-auto bg-[#dc3545] hover:bg-[#c82333] text-white font-bold text-xs px-6 py-2 flex items-center justify-center gap-1.5 cursor-pointer border-0 shadow-sm"
          >
            <span>❌ Thực hiện lại</span>
          </button>

          {/* CÁC TIỆN ÍCH LƯU / THƯ VIỆN QUY TẮC Ở GIỮA */}
          <div className="flex flex-wrap items-center gap-1.5">
            <button
              type="button"
              onClick={handleSaveCurrentAsRule}
              disabled={!expression.trim()}
              className="bg-slate-100 hover:bg-slate-200 disabled:opacity-50 text-slate-800 font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer border border-slate-300"
              title="Lưu biểu thức này vào thư viện quy tắc kiểm tra định kỳ"
            >
              <Save className="w-3.5 h-3.5 text-sky-700" />
              <span>Lưu thành quy tắc</span>
            </button>

            <button
              type="button"
              onClick={() => setShowSavedRulesModal(true)}
              className="bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer border border-slate-300"
              title="Xem và áp dụng các quy tắc đã lưu"
            >
              <Layers className="w-3.5 h-3.5 text-indigo-700" />
              <span>Thư viện quy tắc ({savedRules.length})</span>
            </button>

            <button
              type="button"
              onClick={handleExportResultToExcel}
              className="bg-amber-100 hover:bg-amber-200 text-slate-900 font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer border border-amber-300"
              title="Xuất kết quả tra cứu ra Excel"
            >
              <Download className="w-3.5 h-3.5 text-amber-800" />
              <span>Xuất Excel</span>
            </button>
          </div>

          {/* NÚT TRA CỨU DỮ LIỆU (MÀU XANH LÁ NHƯ ẢNH) */}
          <button
            type="button"
            onClick={handleExecuteQuery}
            className="w-full sm:w-auto bg-[#28a745] hover:bg-[#218838] text-white font-bold text-xs px-8 py-2 flex items-center justify-center gap-1.5 cursor-pointer border-0 shadow-sm"
          >
            <Search className="w-4 h-4 text-white" />
            <span>Tra cứu dữ liệu</span>
          </button>
        </div>
      </div>

      {/* 7. BẢNG DỮ LIỆU BÊN DƯỚI (THỐNG NHẤT 1 KIỂU GIAO DIỆN) */}
      <div className="mx-3 sm:mx-5 mt-4 border border-slate-300 shadow-2xs overflow-hidden">
        {/* THANH TIÊU ĐỀ BẢNG THỐNG NHẤT */}
        <div className="bg-[#286e42] text-white px-3 py-2 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs font-bold">
          <div className="flex items-center gap-2">
            <Table className="w-4 h-4 text-emerald-200" />
            <span>BẢNG DỮ LIỆU PHIẾU ĐIỀU TRA</span>
            {isQueried && (
              <span className="bg-amber-400 text-slate-950 font-bold px-2 py-0.5 text-[11px] rounded-none">
                Đã lọc: {matchedIndices.length.toLocaleString("vi-VN")} / {localData.length.toLocaleString("vi-VN")} dòng (
                {((matchedIndices.length / Math.max(1, localData.length)) * 100).toFixed(1)}%)
              </span>
            )}
          </div>

          {/* CÁC ĐIỀU KHIỂN LỌC NHANH TRONG BẢNG */}
          <div className="flex flex-wrap items-center gap-2">
            {isQueried && (
              <div className="flex items-center bg-white/10 border border-white/20 p-0.5">
                <button
                  type="button"
                  onClick={() => setTableFilterMode("matched_only")}
                  className={`px-2 py-0.5 text-[11px] font-bold cursor-pointer border-0 ${
                    tableFilterMode === "matched_only"
                      ? "bg-white text-emerald-950"
                      : "text-emerald-100 hover:text-white"
                  }`}
                >
                  Chỉ dòng khớp ({matchedIndices.length})
                </button>
                <button
                  type="button"
                  onClick={() => setTableFilterMode("all")}
                  className={`px-2 py-0.5 text-[11px] font-bold cursor-pointer border-0 ${
                    tableFilterMode === "all" ? "bg-white text-emerald-950" : "text-emerald-100 hover:text-white"
                  }`}
                >
                  Tất cả ({localData.length})
                </button>
              </div>
            )}

            {/* Tìm kiếm trong bảng */}
            <div className="relative">
              <input
                type="text"
                value={tableSearchTerm}
                onChange={e => {
                  setTableSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                placeholder="Lọc nhanh bảng..."
                className="bg-white text-slate-900 font-normal text-xs px-2.5 py-1 pr-6 border-0 focus:outline-none w-36 sm:w-44"
              />
              {tableSearchTerm && (
                <button
                  type="button"
                  onClick={() => setTableSearchTerm("")}
                  className="absolute right-1 top-1 text-slate-400 hover:text-slate-800 cursor-pointer border-0 bg-transparent text-xs"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Chọn số dòng mỗi trang */}
            <select
              value={rowsPerPage}
              onChange={e => {
                setRowsPerPage(Number(e.target.value));
                setCurrentPage(1);
              }}
              className="bg-white text-slate-900 font-bold text-xs px-2 py-1 border-0 focus:outline-none cursor-pointer"
            >
              <option value={20}>20 dòng/trang</option>
              <option value={50}>50 dòng/trang</option>
              <option value={100}>100 dòng/trang</option>
              <option value={200}>200 dòng/trang</option>
            </select>
          </div>
        </div>

        {/* BẢNG DỮ LIỆU CUỘN NGANG VÀ DỌC */}
        <div className="max-h-[500px] overflow-auto bg-white">
          <table className="w-full text-left border-collapse text-xs">
            <thead className="bg-slate-100 text-slate-800 sticky top-0 border-b border-slate-300 font-bold z-10 select-none">
              <tr>
                <th className="py-2 px-2.5 w-14 text-center border-r border-slate-200 bg-slate-100">STT</th>
                {isQueried && (
                  <th className="py-2 px-2.5 w-24 text-center border-r border-slate-200 bg-slate-100">
                    TRẠNG THÁI
                  </th>
                )}
                {localColumns.map(col => {
                  const isQueriedCol = usedColumnsInQuery.includes(col);
                  return (
                    <th
                      key={col}
                      className={`py-2 px-3 border-r border-slate-200 whitespace-nowrap ${
                        isQueriedCol ? "bg-amber-100/80 text-amber-950 font-black" : "bg-slate-100 text-slate-800"
                      }`}
                    >
                      {col}
                      {isQueriedCol && <span className="ml-1 text-[10px] text-amber-800 font-normal">(Đang xét)</span>}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {displayedRows.length === 0 ? (
                <tr>
                  <td
                    colSpan={localColumns.length + (isQueried ? 2 : 1)}
                    className="py-12 text-center text-slate-500 italic bg-slate-50/50"
                  >
                    <CheckCircle2 className="w-8 h-8 text-emerald-600 mx-auto mb-1.5" />
                    Không tìm thấy dòng dữ liệu nào khớp với điều kiện tra cứu.
                  </td>
                </tr>
              ) : (
                paginatedIndices.map((rowIdx, displayIdx) => {
                  const row = localData[rowIdx];
                  const isMatched = isQueried && matchedIndices.includes(rowIdx);
                  const stt = (currentPage - 1) * rowsPerPage + displayIdx + 1;

                  return (
                    <tr
                      key={rowIdx}
                      className={`transition-colors ${
                        isMatched
                          ? "bg-amber-50/70 hover:bg-amber-100/70"
                          : displayIdx % 2 === 1
                          ? "bg-slate-50/40 hover:bg-slate-100/50"
                          : "bg-white hover:bg-slate-50"
                      }`}
                    >
                      <td className="py-1.5 px-2.5 text-center text-slate-500 font-mono text-[11px] border-r border-slate-100 font-bold">
                        {stt}
                      </td>

                      {isQueried && (
                        <td className="py-1.5 px-2 text-center border-r border-slate-100">
                          {isMatched ? (
                            <span className="inline-block px-1.5 py-0.5 text-[10px] font-bold bg-rose-100 text-rose-800 border border-rose-300">
                              KHỚP ĐK
                            </span>
                          ) : (
                            <span className="inline-block px-1.5 py-0.5 text-[10px] font-medium text-slate-400">
                              Bình thường
                            </span>
                          )}
                        </td>
                      )}

                      {localColumns.map(col => {
                        const val = row[col];
                        const isQueriedCol = usedColumnsInQuery.includes(col);
                        const isNumber = typeof val === "number" || (!isNaN(Number(val)) && String(val).trim() !== "");

                        return (
                          <td
                            key={col}
                            className={`py-1.5 px-3 border-r border-slate-100 whitespace-nowrap text-slate-800 ${
                              isQueriedCol && isMatched ? "bg-amber-100/50 font-semibold text-amber-950" : ""
                            } ${isNumber ? "font-mono text-right" : ""}`}
                          >
                            {val !== null && val !== undefined && String(val).trim() !== "" ? (
                              isNumber ? (
                                Number(val).toLocaleString("vi-VN")
                              ) : (
                                String(val)
                              )
                            ) : (
                              <span className="text-slate-300 italic text-[11px]">rỗng</span>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {/* THANH PHÂN TRANG BẢNG */}
        <div className="bg-slate-50 px-3 py-2 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-2 text-xs">
          <span className="text-slate-600 font-medium">
            Hiển thị <strong>{paginatedIndices.length}</strong> / <strong>{displayedRows.length.toLocaleString("vi-VN")}</strong> dòng (Trang {currentPage} / {totalPages})
          </span>

          <div className="flex items-center gap-1">
            <button
              type="button"
              disabled={currentPage <= 1}
              onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
              className="bg-white hover:bg-slate-100 disabled:opacity-40 text-slate-700 font-bold px-2 py-1 border border-slate-300 cursor-pointer flex items-center gap-0.5 text-xs"
            >
              <ChevronLeft className="w-3.5 h-3.5" /> Trước
            </button>

            <span className="px-2 py-1 bg-white border border-slate-300 font-bold font-mono text-slate-800">
              {currentPage} / {totalPages}
            </span>

            <button
              type="button"
              disabled={currentPage >= totalPages}
              onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
              className="bg-white hover:bg-slate-100 disabled:opacity-40 text-slate-700 font-bold px-2 py-1 border border-slate-300 cursor-pointer flex items-center gap-0.5 text-xs"
            >
              Sau <ChevronRight className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
      </div>

      {/* 8. MODAL THƯ VIỆN QUY TẮC ĐÃ LƯU */}
      {showSavedRulesModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/40 backdrop-blur-xs flex items-center justify-center p-3">
          <div className="bg-white border border-slate-300 shadow-xl max-w-2xl w-full max-h-[85vh] flex flex-col">
            <div className="bg-[#286e42] text-white px-4 py-2.5 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Layers className="w-4 h-4 text-emerald-200" />
                <span className="font-bold text-sm uppercase">THƯ VIỆN QUY TẮC LOGIC ĐÃ LƯU ({savedRules.length})</span>
              </div>
              <button
                type="button"
                onClick={() => setShowSavedRulesModal(false)}
                className="text-white/80 hover:text-white cursor-pointer border-0 bg-transparent text-sm font-bold"
              >
                ✕
              </button>
            </div>

            <div className="p-3 overflow-y-auto divide-y divide-slate-100 flex-1 space-y-2">
              {savedRules.length === 0 ? (
                <div className="p-6 text-center text-slate-500 italic">
                  Chưa có quy tắc nào được lưu. Bạn có thể soạn biểu thức cho phiếu này và nhấn "Lưu thành quy tắc" để sử dụng lại bất cứ lúc nào.
                </div>
              ) : (
                savedRules.map(rule => (
                  <div key={rule.id} className="pt-2 flex items-start justify-between gap-2">
                    <div className="space-y-0.5">
                      <p className="text-xs font-bold text-slate-900 m-0">{rule.name}</p>
                      <p className="text-xs font-mono text-indigo-700 bg-indigo-50 px-2 py-0.5 border border-indigo-200 inline-block m-0">
                        {rule.expression}
                      </p>
                      <p className="text-[11px] text-slate-500 m-0">{rule.description}</p>
                    </div>

                    <div className="flex items-center gap-1.5 shrink-0">
                      <button
                        type="button"
                        onClick={() => {
                          setExpression(rule.expression);
                          setShowSavedRulesModal(false);
                        }}
                        className="bg-sky-600 hover:bg-sky-700 text-white font-bold text-xs px-2.5 py-1 border-0 cursor-pointer"
                      >
                        Áp dụng
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          if (confirm(`Xóa quy tắc "${rule.name}"?`)) {
                            setSavedRules(prev => prev.filter(r => r.id !== rule.id));
                          }
                        }}
                        className="text-slate-400 hover:text-rose-600 border-0 bg-transparent p-1 cursor-pointer"
                        title="Xóa quy tắc"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>

            <div className="p-3 bg-slate-50 border-t border-slate-200 flex justify-end">
              <button
                type="button"
                onClick={() => setShowSavedRulesModal(false)}
                className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold text-xs px-4 py-1.5 border-0 cursor-pointer"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}
        </div>
      )}
    </div>
  );
};

export default LogicChecking;
