import React, { useState, useMemo } from "react";
import { 
  Database, 
  Sparkles, 
  Search, 
  Terminal, 
  Code2, 
  Table, 
  Download, 
  Play, 
  Zap, 
  Copy, 
  Check, 
  BarChart2, 
  RefreshCw,
  Filter,
  FileSpreadsheet,
  MessageSquareText,
  Lightbulb,
  ArrowRight,
  HelpCircle
} from "lucide-react";
import * as XLSX from "xlsx";

interface ExcelSqlAssistantProps {
  mainData: any[];
  fileName?: string;
}

export const ExcelSqlAssistant: React.FC<ExcelSqlAssistantProps> = ({ mainData, fileName }) => {
  const [queryMode, setQueryMode] = useState<"nl" | "sql" | "formula">("nl");
  const [naturalQuery, setNaturalQuery] = useState("");
  const [sqlQuery, setSqlQuery] = useState("SELECT * FROM data LIMIT 50");
  const [selectedColumn, setSelectedColumn] = useState<string>("");
  const [copied, setCopied] = useState(false);
  const [copiedFormula, setCopiedFormula] = useState(false);
  const [copiedSql, setCopiedSql] = useState(false);
  
  // AI NL-to-SQL states
  const [aiPromptInput, setAiPromptInput] = useState("");
  const [isGeneratingAi, setIsGeneratingAi] = useState(false);
  const [aiResponse, setAiResponse] = useState<{
    sql: string;
    excelFormula: string;
    explanation: string;
  } | null>(null);

  // General AI Dataset Analysis State
  const [aiAnalysis, setAiAnalysis] = useState<string | null>(null);
  const [isAiAnalyzing, setIsAiAnalyzing] = useState(false);

  // Extract columns
  const columns = useMemo(() => {
    if (!mainData || mainData.length === 0) return [];
    return Object.keys(mainData[0] || {});
  }, [mainData]);

  // Suggested Natural Language Prompts based on actual columns
  const samplePrompts = useMemo(() => {
    if (columns.length === 0) return [];
    const col1 = columns[0] || "Tên";
    const col2 = columns.find(c => /doanh thu|sản lượng|giá trị|số|tiền|chi phí/i.test(c)) || columns[1] || "DoanhThu";
    const col3 = columns.find(c => /mã|ngành|vsic|loại|tỉnh|huyện/i.test(c)) || columns[2] || "MaNganh";

    return [
      {
        label: `🔍 Lọc các bản ghi có ${col2} > 1,000,000`,
        prompt: `Lọc danh sách các dòng có chỉ tiêu ${col2} lớn hơn 1000000`
      },
      {
        label: `📊 Đếm số lượng bản ghi nhóm theo ${col3}`,
        prompt: `Đếm số lượng bản ghi phân nhóm theo cột ${col3}`
      },
      {
        label: `⚠️ Lọc các dòng bị thiếu hoặc trống dữ liệu ở ${col1}`,
        prompt: `Tìm các bản ghi bị ô trống hoặc thiếu thông tin tại cột ${col1}`
      },
      {
        label: `📈 Lấy Top 10 dòng có ${col2} cao nhất`,
        prompt: `Sắp xếp giảm dần theo ${col2} và lấy 10 dòng có giá trị lớn nhất`
      }
    ];
  }, [columns]);

  // SQL / NL filter simulation
  const filteredResults = useMemo(() => {
    if (!mainData || mainData.length === 0) return [];

    let result = [...mainData];

    if (queryMode === "nl" && naturalQuery.trim()) {
      const q = naturalQuery.toLowerCase().trim();
      result = result.filter(row => {
        return Object.values(row).some(val => 
          String(val ?? "").toLowerCase().includes(q)
        );
      });
    } else if (queryMode === "sql" && sqlQuery.trim()) {
      // Basic SQL WHERE / LIMIT clause parser simulation
      try {
        const upper = sqlQuery.toUpperCase();
        if (upper.includes("WHERE")) {
          const wherePart = sqlQuery.split(/WHERE/i)[1]?.split(/LIMIT/i)[0]?.trim();
          if (wherePart) {
            const match = wherePart.match(/([a-zA-Z0-9_\u00C0-\u024F\u1E00-\u1EFF\s]+)\s*(=|LIKE|>|<|>=|<=)\s*['"]?([^'"]+)['"]?/i);
            if (match) {
              const [, colName, op, val] = match;
              const actualCol = columns.find(c => c.toLowerCase() === colName.trim().toLowerCase()) || columns[0];
              if (actualCol) {
                result = result.filter(row => {
                  const cellVal = String(row[actualCol] ?? "");
                  const targetVal = val.replace(/%/g, "").trim();
                  if (op === "=") return cellVal.toLowerCase() === targetVal.toLowerCase();
                  if (op.toUpperCase() === "LIKE") return cellVal.toLowerCase().includes(targetVal.toLowerCase());
                  if (op === ">") return Number(cellVal.replace(/,/g, "")) > Number(targetVal.replace(/,/g, ""));
                  if (op === "<") return Number(cellVal.replace(/,/g, "")) < Number(targetVal.replace(/,/g, ""));
                  if (op === ">=") return Number(cellVal.replace(/,/g, "")) >= Number(targetVal.replace(/,/g, ""));
                  if (op === "<=") return Number(cellVal.replace(/,/g, "")) <= Number(targetVal.replace(/,/g, ""));
                  return true;
                });
              }
            }
          }
        }
        if (upper.includes("LIMIT")) {
          const limitNum = parseInt(sqlQuery.split(/LIMIT/i)[1]?.trim() || "50", 10);
          if (!isNaN(limitNum) && limitNum > 0) {
            result = result.slice(0, limitNum);
          }
        }
      } catch (e) {
        // Fallback to full result
      }
    }

    return result;
  }, [mainData, queryMode, naturalQuery, sqlQuery, columns]);

  // Column statistics calculation
  const colStats = useMemo(() => {
    if (!selectedColumn || !mainData || mainData.length === 0) return null;
    const values = mainData
      .map(row => row[selectedColumn])
      .filter(val => val !== null && val !== undefined && val !== "");

    const numericValues = values
      .map(v => Number(String(v).replace(/,/g, "")))
      .filter(n => !isNaN(n));

    const totalCount = values.length;
    const isNumeric = numericValues.length > totalCount * 0.5;

    if (isNumeric && numericValues.length > 0) {
      const sum = numericValues.reduce((acc, curr) => acc + curr, 0);
      const avg = sum / numericValues.length;
      const min = Math.min(...numericValues);
      const max = Math.max(...numericValues);
      return {
        type: "numeric",
        count: totalCount,
        sum: sum.toLocaleString("vi-VN"),
        avg: avg.toLocaleString("vi-VN", { maximumFractionDigits: 2 }),
        min: min.toLocaleString("vi-VN"),
        max: max.toLocaleString("vi-VN")
      };
    } else {
      const freqMap: Record<string, number> = {};
      values.forEach(v => {
        const key = String(v).trim();
        freqMap[key] = (freqMap[key] || 0) + 1;
      });
      const uniqueCount = Object.keys(freqMap).length;
      const topValues = Object.entries(freqMap)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

      return {
        type: "text",
        count: totalCount,
        unique: uniqueCount,
        topValues
      };
    }
  }, [selectedColumn, mainData]);

  // Handle Export Filtered
  const handleExportFiltered = () => {
    if (filteredResults.length === 0) return;
    const ws = XLSX.utils.json_to_sheet(filteredResults);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SQL_Result");
    XLSX.writeFile(wb, `Ket_Qua_Truy_Van_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // AI Gemini Convert NL to SQL & Excel Formula
  const handleGenerateSqlFromNl = async (customPrompt?: string) => {
    const promptToUse = customPrompt || aiPromptInput;
    if (!promptToUse.trim()) return;

    setIsGeneratingAi(true);
    setAiResponse(null);

    try {
      const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || (process.env as any)?.GEMINI_API_KEY;
      if (!apiKey) {
        setAiResponse({
          sql: `-- Lỗi: Chưa cấu hình VITE_GEMINI_API_KEY`,
          excelFormula: `=FILTER(A2:Z1000, A2:A1000>0)`,
          explanation: "⚠️ Chưa có VITE_GEMINI_API_KEY trong .env. Vui lòng kiểm tra lại chìa khóa AI Gemini."
        });
        setIsGeneratingAi(false);
        return;
      }

      const promptText = `Bạn là chuyên gia cơ sở dữ liệu SQL và công thức Excel. 
Người dùng muốn thực hiện yêu cầu sau trên bảng dữ liệu có tên là "data":
Yêu cầu bằng ngôn ngữ tự nhiên: "${promptToUse}"

Danh sách các cột trong bảng dữ liệu thực tế: [${columns.join(", ")}].

Hãy phản hồi theo định dạng JSON chuẩn (chỉ trả về JSON, không kèm Markdown code fence):
{
  "sql": "Câu lệnh SQL chuẩn với tên cột thực tế (ví dụ: SELECT * FROM data WHERE ColName LIKE '%val%' LIMIT 50)",
  "excelFormula": "Công thức Excel tương ứng (ví dụ: =FILTER(...) hoặc =COUNTIF(...) hoặc =SUMIFS(...))",
  "explanation": "Giải thích ngắn gọn 1-2 câu tiếng Việt về cách lệnh này hoạt động"
}`;

      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: promptText }] }]
        })
      });

      const data = await response.json();
      const rawText = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
      const cleanedJson = rawText.replace(/```json/g, "").replace(/```/g, "").trim();

      try {
        const parsed = JSON.parse(cleanedJson);
        setAiResponse({
          sql: parsed.sql || `SELECT * FROM data LIMIT 50`,
          excelFormula: parsed.excelFormula || `=FILTER(A:Z, A:A<>"")`,
          explanation: parsed.explanation || "Đã sinh câu lệnh dựa trên các cột thực tế của bảng."
        });
        if (parsed.sql) {
          setSqlQuery(parsed.sql);
        }
      } catch (e) {
        setAiResponse({
          sql: `SELECT * FROM data WHERE ${columns[0] || "col"} LIKE '%${promptToUse}%'`,
          excelFormula: `=FILTER(data, SEARCH("${promptToUse}", A:A))`,
          explanation: rawText || "Đã sinh câu lệnh truy vấn mẫu."
        });
      }
    } catch (err: any) {
      setAiResponse({
        sql: `SELECT * FROM data LIMIT 50`,
        excelFormula: `=FILTER(...)`,
        explanation: `Lỗi kết nối AI: ${err.message || err}`
      });
    } finally {
      setIsGeneratingAi(false);
    }
  };

  // AI Gemini General Analysis Trigger
  const handleAiAnalyze = async () => {
    setIsAiAnalyzing(true);
    setAiAnalysis(null);

    try {
      const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || (process.env as any)?.GEMINI_API_KEY;
      if (!apiKey) {
        setAiAnalysis("⚠️ Chưa cấu hình VITE_GEMINI_API_KEY trong tệp .env. Vui lòng kiểm tra cấu hình.");
        setIsAiAnalyzing(false);
        return;
      }

      const sampleRows = mainData.slice(0, 15);
      const promptText = `Bạn là trợ lý chuyên gia dữ liệu Excel và SQL. Hãy phân tích tập dữ liệu ${fileName || "hiện tại"} có ${mainData.length} dòng và các cột [${columns.join(", ")}].
Một số dòng mẫu:
${JSON.stringify(sampleRows, null, 2)}

Hãy đưa ra tóm tắt phân tích ngắn gọn, súc tích bằng tiếng Việt gồm:
1. Nhận xét tổng quan cấu trúc dữ liệu.
2. Đề xuất 3 câu lệnh SQL hoặc bộ lọc Excel hữu ích để đào sâu dữ liệu này.
3. Cảnh báo sai sót dữ liệu nếu có (ví dụ: dòng trống, dữ liệu không nhất quán).`;

      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: promptText }] }]
        })
      });

      const data = await response.json();
      const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (text) {
        setAiAnalysis(text);
      } else {
        setAiAnalysis("Không thể kết nối dịch vụ AI. Vui lòng thử lại sau.");
      }
    } catch (err: any) {
      setAiAnalysis(`Lỗi kết nối AI: ${err.message || err}`);
    } finally {
      setIsAiAnalyzing(false);
    }
  };

  const copyText = (text: string, type: "general" | "sql" | "formula") => {
    navigator.clipboard.writeText(text);
    if (type === "sql") {
      setCopiedSql(true);
      setTimeout(() => setCopiedSql(false), 2000);
    } else if (type === "formula") {
      setCopiedFormula(true);
      setTimeout(() => setCopiedFormula(false), 2000);
    } else {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  return (
    <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-sm">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4">
        <div className="flex items-center gap-3">
          <div className="p-2.5 bg-indigo-50 border border-indigo-200 rounded-xl">
            <Database className="w-6 h-6 text-indigo-600" />
          </div>
          <div>
            <h2 className="text-base font-extrabold text-slate-900 flex items-center gap-2">
              TRỢ LÝ EXCEL & TRUY VẤN SQL THÔNG MINH
              <span className="text-[10px] bg-indigo-100 text-indigo-800 font-extrabold px-2.5 py-0.5 rounded-full border border-indigo-200 uppercase">
                AI Powered
              </span>
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Hỏi bằng ngôn ngữ tự nhiên để AI tự động sinh câu lệnh SQL, công thức Excel & chạy lọc dữ liệu trực tiếp.
            </p>
          </div>
        </div>

        <button
          onClick={handleAiAnalyze}
          disabled={isAiAnalyzing || mainData.length === 0}
          className="bg-gradient-to-r from-indigo-600 to-sky-600 hover:from-indigo-700 hover:to-sky-700 text-white font-bold text-xs px-4 py-2.5 rounded-xl shadow-sm transition-all cursor-pointer flex items-center justify-center gap-2 disabled:opacity-50"
        >
          {isAiAnalyzing ? (
            <>
              <RefreshCw className="w-4 h-4 animate-spin text-white" /> Đang Phân Tích AI...
            </>
          ) : (
            <>
              <Sparkles className="w-4 h-4 text-amber-300 fill-amber-300" /> AI Phân Tích Toàn Bảng
            </>
          )}
        </button>
      </div>

      {/* AI Insights Card if generated */}
      {aiAnalysis && (
        <div className="bg-gradient-to-r from-indigo-50/80 via-sky-50/80 to-white border border-indigo-200 rounded-2xl p-5 space-y-3 animate-fade-in relative">
          <div className="flex items-center justify-between border-b border-indigo-150 pb-2">
            <div className="flex items-center gap-2 text-indigo-900 font-extrabold text-xs">
              <Sparkles className="w-4 h-4 text-indigo-600 fill-indigo-200" />
              KẾT QUẢ PHÂN TÍCH AI GEMINI VỀ DỮ LIỆU
            </div>
            <button 
              onClick={() => copyText(aiAnalysis, "general")}
              className="text-xs text-indigo-700 hover:text-indigo-900 flex items-center gap-1 font-bold cursor-pointer"
            >
              {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
              {copied ? "Đã chép" : "Sao chép"}
            </button>
          </div>
          <div className="text-xs text-slate-700 leading-relaxed whitespace-pre-wrap font-sans">
            {aiAnalysis}
          </div>
        </div>
      )}

      {/* AI Prompt generator Section (HIGHLIGHT FEATURE) */}
      <div className="bg-gradient-to-br from-indigo-900 via-slate-900 to-indigo-950 text-white rounded-2xl p-5 space-y-4 shadow-md border border-indigo-500/30">
        <div className="flex items-center justify-between border-b border-white/10 pb-3">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-indigo-500/20 border border-indigo-400/40 flex items-center justify-center text-sky-300">
              <Sparkles className="w-4 h-4 text-amber-300 fill-amber-300" />
            </div>
            <div>
              <h3 className="text-sm font-extrabold text-white flex items-center gap-2">
                AI CHUYỂN NGÔN NGỮ TỰ NHIÊN THÀNH LỆNH SQL & CÔNG THỨC EXCEL
              </h3>
              <p className="text-[11.5px] text-slate-300">
                Gõ câu hỏi bất kỳ hoặc chọn ví dụ mẫu, AI Gemini sẽ đọc cột thực tế & sinh câu lệnh phù hợp.
              </p>
            </div>
          </div>
        </div>

        {/* Input & Action */}
        <div className="space-y-3">
          <div className="flex flex-col sm:flex-row gap-2">
            <input
              type="text"
              value={aiPromptInput}
              onChange={(e) => setAiPromptInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") handleGenerateSqlFromNl(); }}
              placeholder="Nhập yêu cầu bằng tiếng Việt (ví dụ: Lọc các đơn vị thuộc Hà Nội có doanh thu > 5 tỷ)..."
              className="flex-1 bg-slate-950/80 border border-indigo-500/40 rounded-xl px-4 py-2.5 text-xs text-white placeholder:text-slate-400 focus:outline-none focus:border-sky-400 font-sans shadow-inner"
            />
            <button
              onClick={() => handleGenerateSqlFromNl()}
              disabled={isGeneratingAi || !aiPromptInput.trim()}
              className="bg-gradient-to-r from-indigo-500 to-sky-500 hover:from-indigo-600 hover:to-sky-600 text-white font-black text-xs px-5 py-2.5 rounded-xl shadow-md transition-all cursor-pointer flex items-center justify-center gap-2 shrink-0 disabled:opacity-50"
            >
              {isGeneratingAi ? (
                <>
                  <RefreshCw className="w-4 h-4 animate-spin text-white" /> AI Đang Sinh Lệnh...
                </>
              ) : (
                <>
                  <Play className="w-4 h-4 fill-current text-amber-300" /> AI Tạo Lệnh &amp; Chạy
                </>
              )}
            </button>
          </div>

          {/* Example Prompt Chips */}
          {samplePrompts.length > 0 && (
            <div className="space-y-1.5 pt-1">
              <div className="text-[11px] font-bold text-indigo-300 flex items-center gap-1">
                <Lightbulb className="w-3.5 h-3.5 text-amber-400" /> Mẫu câu hỏi tự nhiên gợi ý (bấm để thử ngay):
              </div>
              <div className="flex flex-wrap gap-2">
                {samplePrompts.map((item, index) => (
                  <button
                    key={index}
                    onClick={() => {
                      setAiPromptInput(item.prompt);
                      handleGenerateSqlFromNl(item.prompt);
                    }}
                    className="bg-white/10 hover:bg-white/20 border border-white/15 hover:border-sky-400/50 text-slate-200 text-[11px] font-medium px-3 py-1.5 rounded-lg transition-all cursor-pointer flex items-center gap-1.5 active:scale-95 text-left"
                  >
                    <span>{item.label}</span>
                    <ArrowRight className="w-3 h-3 text-sky-300 shrink-0" />
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* AI Response Output Card */}
        {aiResponse && (
          <div className="bg-slate-950/90 border border-indigo-500/40 rounded-xl p-4 space-y-3 animate-fade-in text-xs">
            <div className="flex items-center justify-between border-b border-white/10 pb-2">
              <span className="font-extrabold text-sky-300 flex items-center gap-1.5">
                <Check className="w-4 h-4 text-emerald-400" /> KẾT QUẢ AI TẠO CÂU LỆNH MẪU
              </span>
              <span className="text-[10px] text-slate-400 italic">{aiResponse.explanation}</span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              {/* Generated SQL */}
              <div className="bg-slate-900 border border-slate-800 rounded-lg p-3 space-y-1.5">
                <div className="flex items-center justify-between text-[11px] font-bold text-indigo-300">
                  <span className="flex items-center gap-1"><Terminal className="w-3.5 h-3.5" /> Câu Lệnh SQL:</span>
                  <button
                    onClick={() => copyText(aiResponse.sql, "sql")}
                    className="text-[10px] text-slate-400 hover:text-white flex items-center gap-1 cursor-pointer"
                  >
                    {copiedSql ? <Check className="w-3 h-3 text-emerald-400" /> : <Copy className="w-3 h-3" />}
                    {copiedSql ? "Đã chép" : "Sao chép"}
                  </button>
                </div>
                <code className="block bg-black/60 p-2 rounded text-emerald-400 font-mono text-[11px] break-all">
                  {aiResponse.sql}
                </code>
                <button
                  onClick={() => {
                    setQueryMode("sql");
                    setSqlQuery(aiResponse.sql);
                  }}
                  className="mt-1 w-full bg-indigo-600/80 hover:bg-indigo-600 text-white font-bold text-[10.5px] py-1 rounded transition-colors cursor-pointer flex items-center justify-center gap-1"
                >
                  <Play className="w-3 h-3 fill-current" /> Chạy Câu Lệnh SQL Này
                </button>
              </div>

              {/* Generated Excel Formula */}
              <div className="bg-slate-900 border border-slate-800 rounded-lg p-3 space-y-1.5">
                <div className="flex items-center justify-between text-[11px] font-bold text-emerald-300">
                  <span className="flex items-center gap-1"><FileSpreadsheet className="w-3.5 h-3.5" /> Công Thức Excel Tương Ứng:</span>
                  <button
                    onClick={() => copyText(aiResponse.excelFormula, "formula")}
                    className="text-[10px] text-slate-400 hover:text-white flex items-center gap-1 cursor-pointer"
                  >
                    {copiedFormula ? <Check className="w-3 h-3 text-emerald-400" /> : <Copy className="w-3 h-3" />}
                    {copiedFormula ? "Đã chép" : "Sao chép"}
                  </button>
                </div>
                <code className="block bg-black/60 p-2 rounded text-sky-300 font-mono text-[11px] break-all">
                  {aiResponse.excelFormula}
                </code>
                <span className="text-[10px] text-slate-400 block pt-1">
                  💡 Copy công thức này dán trực tiếp vào ô Excel của bạn.
                </span>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Mode Switcher */}
      <div className="flex items-center gap-2 border-b border-slate-200 pb-3">
        <button
          onClick={() => setQueryMode("nl")}
          className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 cursor-pointer ${
            queryMode === "nl"
              ? "bg-indigo-600 text-white shadow-sm"
              : "bg-slate-100 text-slate-700 hover:bg-slate-200"
          }`}
        >
          <Search className="w-3.5 h-3.5" /> Lọc Ngôn Ngữ Tự Nhiên
        </button>
        <button
          onClick={() => setQueryMode("sql")}
          className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 cursor-pointer ${
            queryMode === "sql"
              ? "bg-indigo-600 text-white shadow-sm"
              : "bg-slate-100 text-slate-700 hover:bg-slate-200"
          }`}
        >
          <Terminal className="w-3.5 h-3.5" /> Trình Soạn Thảo &amp; Chạy Lệnh SQL
        </button>
        <button
          onClick={() => setQueryMode("formula")}
          className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 cursor-pointer ${
            queryMode === "formula"
              ? "bg-indigo-600 text-white shadow-sm"
              : "bg-slate-100 text-slate-700 hover:bg-slate-200"
          }`}
        >
          <BarChart2 className="w-3.5 h-3.5" /> Phân Tích Thống Kê Cột
        </button>
      </div>

      {/* Mode 1: Natural Language Filter */}
      {queryMode === "nl" && (
        <div className="space-y-3 bg-slate-50 border border-slate-200 p-4 rounded-xl">
          <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
            <Filter className="w-3.5 h-3.5 text-indigo-600" /> Nhập từ khóa hoặc cụm từ tìm kiếm trong toàn bộ bảng:
          </label>
          <div className="flex gap-2">
            <input
              type="text"
              value={naturalQuery}
              onChange={(e) => setNaturalQuery(e.target.value)}
              placeholder="Ví dụ: Nông nghiệp, Xã Vĩnh Lộc, 12345,..."
              className="flex-1 bg-white border border-slate-300 rounded-xl px-3.5 py-2 text-xs text-slate-800 focus:outline-none focus:border-indigo-500 font-sans shadow-2xs"
            />
            {naturalQuery && (
              <button
                onClick={() => setNaturalQuery("")}
                className="px-3 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 text-xs font-bold rounded-xl transition-all cursor-pointer"
              >
                Xóa
              </button>
            )}
          </div>
          <p className="text-[11px] text-slate-500">
            * Hệ thống tự động quét và lọc tất cả các dòng chứa cụm từ tìm kiếm trên bất kỳ cột nào.
          </p>
        </div>
      )}

      {/* Mode 2: SQL Query Editor */}
      {queryMode === "sql" && (
        <div className="space-y-3 bg-slate-900 border border-slate-800 p-4 rounded-xl text-white">
          <div className="flex items-center justify-between">
            <label className="text-xs font-mono font-bold text-indigo-300 flex items-center gap-1.5">
              <Code2 className="w-3.5 h-3.5" /> NHẬP CÂU LỆNH TRUY VẤN SQL CỦA BẠN:
            </label>
            <span className="text-[10px] text-slate-400 font-mono">Hỗ trợ WHERE, LIKE, LIMIT</span>
          </div>
          <textarea
            value={sqlQuery}
            onChange={(e) => setSqlQuery(e.target.value)}
            rows={3}
            className="w-full bg-slate-950 border border-slate-800 rounded-lg p-3 text-xs font-mono text-emerald-400 focus:outline-none focus:border-indigo-500 shadow-inner"
            placeholder="SELECT * FROM data WHERE Tên_Cột LIKE '%giá_trị%' LIMIT 50"
          />
          <div className="flex items-center justify-between pt-1 text-[11px] text-slate-400">
            <div className="flex items-center gap-1.5 flex-wrap">
              <span className="text-slate-500">Các cột thực tế trong bảng:</span>
              {columns.slice(0, 5).map(col => (
                <button
                  key={col}
                  onClick={() => setSqlQuery(`SELECT * FROM data WHERE ${col} LIKE '%A%' LIMIT 50`)}
                  className="bg-slate-800 hover:bg-slate-700 text-indigo-300 px-2 py-0.5 rounded text-[10px] font-mono cursor-pointer"
                >
                  {col}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Mode 3: Formula & Column Stats */}
      {queryMode === "formula" && (
        <div className="space-y-4 bg-slate-50 border border-slate-200 p-4 rounded-xl">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="text-xs font-bold text-slate-700 block mb-1.5">
                Chọn Cột Để Xem Thống Kê & Nhanh:
              </label>
              <select
                value={selectedColumn}
                onChange={(e) => setSelectedColumn(e.target.value)}
                className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-indigo-500 font-sans shadow-2xs"
              >
                <option value="">-- Chọn một cột --</option>
                {columns.map(col => (
                  <option key={col} value={col}>{col}</option>
                ))}
              </select>
            </div>

            {colStats && (
              <div className="bg-white border border-slate-200 p-3 rounded-xl space-y-2">
                <div className="text-xs font-bold text-indigo-900 border-b border-slate-100 pb-1 flex items-center gap-1">
                  <Zap className="w-3.5 h-3.5 text-indigo-600" />
                  KẾT QUẢ THỐNG KÊ CỘT: <span className="text-slate-800">{selectedColumn}</span>
                </div>
                {colStats.type === "numeric" ? (
                  <div className="grid grid-cols-2 gap-2 text-[11px]">
                    <div className="bg-slate-50 p-1.5 rounded"><span className="text-slate-500">Tổng số dòng:</span> <b className="text-slate-800">{colStats.count}</b></div>
                    <div className="bg-slate-50 p-1.5 rounded"><span className="text-slate-500">Tổng SUM:</span> <b className="text-indigo-700">{colStats.sum}</b></div>
                    <div className="bg-slate-50 p-1.5 rounded"><span className="text-slate-500">Trung bình (AVG):</span> <b className="text-slate-800">{colStats.avg}</b></div>
                    <div className="bg-slate-50 p-1.5 rounded"><span className="text-slate-500">Nhỏ / Lớn nhất:</span> <b className="text-slate-800">{colStats.min} / {colStats.max}</b></div>
                  </div>
                ) : (
                  <div className="space-y-1.5 text-[11px]">
                    <div className="flex justify-between text-slate-600">
                      <span>Số dòng có giá trị: <b>{colStats.count}</b></span>
                      <span>Số giá trị khác nhau: <b>{colStats.unique}</b></span>
                    </div>
                    <div className="text-[10px] font-bold text-slate-500 uppercase mt-1">Top giá trị xuất hiện nhiều nhất:</div>
                    <div className="space-y-1">
                      {colStats.topValues?.map(([val, cnt]) => (
                        <div key={val} className="flex justify-between bg-slate-50 px-2 py-0.5 rounded text-[10.5px]">
                          <span className="truncate max-w-[200px] text-slate-800 font-medium">{val || "(Trống)"}</span>
                          <span className="font-mono text-indigo-600 font-bold">{cnt} lần</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Results Table Section */}
      <div className="space-y-3">
        <div className="flex items-center justify-between">
          <div className="text-xs font-bold text-slate-800 flex items-center gap-2">
            <Table className="w-4 h-4 text-indigo-600" />
            KẾT QUẢ TRUY VẤN DỮ LIỆU ({filteredResults.length.toLocaleString("vi-VN")} DÒNG)
          </div>
          <button
            onClick={handleExportFiltered}
            disabled={filteredResults.length === 0}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-3.5 py-1.5 rounded-xl transition-all cursor-pointer flex items-center gap-1.5 disabled:opacity-50"
          >
            <Download className="w-3.5 h-3.5" /> Xuất Excel Kết Quả
          </button>
        </div>

        {filteredResults.length === 0 ? (
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-8 text-center text-xs text-slate-500">
            Không tìm thấy dữ liệu phù hợp với điều kiện truy vấn.
          </div>
        ) : (
          <div className="border border-slate-200 rounded-xl overflow-x-auto max-h-96 custom-scrollbar">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-slate-100 text-slate-700 font-bold sticky top-0 border-b border-slate-200">
                <tr>
                  <th className="p-2.5 border-r border-slate-200 w-12 text-center">STT</th>
                  {columns.map((col) => (
                    <th key={col} className="p-2.5 border-r border-slate-200 whitespace-nowrap min-w-[120px]">
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 text-slate-800">
                {filteredResults.slice(0, 100).map((row, idx) => (
                  <tr key={idx} className="hover:bg-indigo-50/40 transition-colors">
                    <td className="p-2 text-center text-slate-500 font-mono border-r border-slate-200 bg-slate-50/50">
                      {idx + 1}
                    </td>
                    {columns.map((col) => (
                      <td key={col} className="p-2 border-r border-slate-200 truncate max-w-xs">
                        {String(row[col] ?? "")}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
            {filteredResults.length > 100 && (
              <div className="p-2 bg-slate-50 text-center text-[11px] text-slate-500 border-t border-slate-200 font-medium">
                Đang hiển thị 100 / {filteredResults.length.toLocaleString("vi-VN")} dòng đầu tiên. Nhấn nút "Xuất Excel Kết Quả" để tải về toàn bộ.
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default ExcelSqlAssistant;
