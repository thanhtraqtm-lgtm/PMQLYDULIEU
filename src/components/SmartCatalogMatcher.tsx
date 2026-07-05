import React, { useState, useEffect, useMemo } from "react";
import { 
  CheckSquare, 
  AlertTriangle, 
  Download, 
  Search, 
  RefreshCw, 
  Zap, 
  Brain, 
  Sliders, 
  Database,
  ArrowRightLeft,
  CheckCircle2,
  Info
} from "lucide-react";
import { GoogleGenAI } from "@google/genai";
import * as XLSX from "xlsx";
import { vsicRawData } from "../data/vsic";
import { LiveTenRowPreview } from "./LiveTenRowPreview";

interface ColumnMapping {
  mota: string;
  manganh: string;
  xa: string;
  doanhthu: string;
  laodong: string;
  idCol: string;
}

interface SmartCatalogMatcherProps {
  mainData: any[];
  columns: string[];
  mapping: ColumnMapping;
  onUpdateMainData: (newData: any[]) => void;
  onExportExcel?: (data: any[], fileName: string) => void;
}

// Vietnamese standard character cleaning
function cleanVietnameseString(str: string): string {
  if (!str) return "";
  return str.toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[đĐ]/g, "d")
    .replace(/[^a-z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Calculate token-level overlap Jaccard index
function getWordJaccardSimilarity(str1: string, str2: string): number {
  const words1 = new Set(cleanVietnameseString(str1).split(" ").filter(Boolean));
  const words2 = new Set(cleanVietnameseString(str2).split(" ").filter(Boolean));
  if (words1.size === 0 || words2.size === 0) return 0;
  
  let intersectionSize = 0;
  words1.forEach(w => {
    if (words2.has(w)) intersectionSize++;
  });
  
  const unionSize = words1.size + words2.size - intersectionSize;
  return intersectionSize / unionSize;
}

export default function SmartCatalogMatcher({
  mainData,
  columns,
  mapping,
  onUpdateMainData,
  onExportExcel
}: SmartCatalogMatcherProps) {
  // Config states
  const [sourceCol, setSourceCol] = useState<string>("");
  const [catalogType, setCatalogType] = useState<"vsic" | "custom">("vsic");
  const [customCatalogText, setCustomCatalogText] = useState<string>(
    "Mã, Tên danh mục\n101, Chăn nuôi heo/lợn thịt\n102, Trồng trọt lúa gạo\n201, Khai thác thủy sản ven bờ\n301, Bán lẻ tạp hóa bách hóa tổng hợp"
  );
  
  const [useAi, setUseAi] = useState<boolean>(true);
  const [aiThreshold, setAiThreshold] = useState<number>(75); // Use AI if fuzzy confidence is below 75%
  const [batchLimit, setBatchLimit] = useState<string>("50"); // "10", "50", "200", "all"
  const [targetColName, setTargetColName] = useState<string>("Ma_Nganh_Doi_Sanh");
  const [targetNameColName, setTargetNameColName] = useState<string>("Ten_Nganh_Doi_Sanh");

  // Run states
  const [isMatching, setIsMatching] = useState<boolean>(false);
  const [progress, setProgress] = useState<number>(0);
  const [currentAction, setCurrentAction] = useState<string>("");
  const [matchResults, setMatchResults] = useState<any[]>([]);
  const [showResults, setShowResults] = useState<boolean>(false);
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [filterMode, setFilterMode] = useState<"all" | "high" | "low" | "unmatched">("all");
  const [catalogMatchStats, setCatalogMatchStats] = useState<{ total: number; matched: number; unmatched: number; avgConfidence: number } | null>(null);

  // Load standard VSIC candidates once
  const vsicCatalogList = useMemo(() => {
    const list: { code: string; name: string; cleanName: string }[] = [];
    if (vsicRawData) {
      Object.entries(vsicRawData).forEach(([code, name]) => {
        // Only focus on 5-digit codes for high-precision catalog matching, or all if short
        if (code.length === 5 || code.length === 4) {
          list.push({
            code,
            name,
            cleanName: cleanVietnameseString(name)
          });
        }
      });
    }
    return list;
  }, []);

  // Parse pasted custom catalog text
  const customCatalogList = useMemo(() => {
    const list: { code: string; name: string; cleanName: string }[] = [];
    if (catalogType === "custom" && customCatalogText) {
      const lines = customCatalogText.split("\n");
      lines.forEach((line, i) => {
        if (i === 0 && (line.includes("Mã") || line.includes("Name") || line.includes("Code"))) return; // Skip header
        const parts = line.split(/[,;\t]/);
        if (parts.length >= 2) {
          const code = parts[0].trim();
          const name = parts.slice(1).join(",").trim().replace(/^["']|["']$/g, "");
          if (code && name) {
            list.push({
              code,
              name,
              cleanName: cleanVietnameseString(name)
            });
          }
        }
      });
    }
    return list;
  }, [catalogType, customCatalogText]);

  const activeCatalog = catalogType === "vsic" ? vsicCatalogList : customCatalogList;

  // Auto-select source column without overwriting active user selection
  useEffect(() => {
    if (columns && columns.length > 0) {
      setSourceCol(prev => {
        if (prev && columns.includes(prev)) return prev;
        return mapping.mota || columns.find(c => {
          const lower = c.toLowerCase();
          return lower.includes("mô tả") || lower.includes("mota") || lower.includes("hoạt động") || lower.includes("kinh doanh") || lower.includes("tên hộ");
        }) || columns[0];
      });
    }
  }, [columns, mapping]);

  // Jaccard similarity and substring booster match function
  const getFuzzySuggestions = (desc: string, catalog: typeof activeCatalog, limit = 5) => {
    if (!desc) return [];
    const cleanedDesc = cleanVietnameseString(desc);
    
    const scores = catalog.map(item => {
      const jaccard = getWordJaccardSimilarity(desc, item.name);
      
      // Check if standard category words are found inside the actual description (substring check)
      let substringBonus = 0;
      if (cleanedDesc.includes(item.cleanName) || item.cleanName.includes(cleanedDesc)) {
        substringBonus = 0.3;
      }

      // Exact phrase match bonus
      let phraseBonus = 0;
      const descTokens = cleanedDesc.split(" ");
      const nameTokens = item.cleanName.split(" ");
      if (descTokens.length > 0 && nameTokens.length > 0) {
        if (cleanedDesc.includes(item.cleanName)) phraseBonus = 0.2;
      }

      const finalScore = Math.min(1.0, jaccard * 0.6 + substringBonus + phraseBonus);

      return {
        item,
        score: Math.round(finalScore * 100)
      };
    });

    // Sort by score descending
    scores.sort((a, b) => b.score - a.score);
    return scores.slice(0, limit);
  };

  // Run matching batch job (Fuzzy + optional client-side Gemini)
  const handleStartMatch = async () => {
    if (mainData.length === 0) return;
    if (!sourceCol) {
      alert("Vui lòng chọn cột mô tả dữ liệu nguồn!");
      return;
    }
    if (activeCatalog.length === 0) {
      alert("Danh mục so sánh trống hoặc chưa parse được! Hãy kiểm tra lại.");
      return;
    }

    const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || "";
    if (useAi && !apiKey) {
      alert("Bạn đã bật chế độ AI Gemini nhưng chưa cấu hình VITE_GEMINI_API_KEY trong Cài đặt > Khóa bí mật (Secrets) hoặc file .env.\n\nHệ thống sẽ tự động tắt chế độ AI nâng cao và chạy hoàn toàn bằng bộ lọc so khớp mờ local.");
      setUseAi(false);
    }

    setIsMatching(true);
    setProgress(0);
    setMatchResults([]);

    // Determine row range to scan
    let maxRows = mainData.length;
    if (batchLimit !== "all") {
      maxRows = Math.min(mainData.length, parseInt(batchLimit, 10));
    }

    const tempResults: any[] = [];
    let aiClient: any = null;

    if (useAi && apiKey) {
      aiClient = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build"
          }
        }
      });
    }

    for (let i = 0; i < maxRows; i++) {
      const row = mainData[i];
      const descValue = String(row[sourceCol] || "").trim();

      if (!descValue) {
        tempResults.push({
          rowIdx: i,
          rowNum: i + 1,
          originalDesc: "[RỖNG]",
          matchedCode: "",
          matchedName: "Không có dữ liệu mô tả",
          confidence: 0,
          method: "Bỏ qua",
          reason: "Dòng rỗng",
          suggestions: []
        });
        continue;
      }

      setCurrentAction(`Đang so khớp dòng ${i + 1}/${maxRows}: "${descValue.substring(0, 30)}..."`);

      // 1. Local Fuzzy similarity
      const topSuggestions = getFuzzySuggestions(descValue, activeCatalog, 5);
      const topMatch = topSuggestions[0] || null;

      let matchedCode = "";
      let matchedName = "";
      let confidence = 0;
      let method = "So khớp từ vựng";
      let reason = "";

      if (topMatch && topMatch.score >= 30) {
        matchedCode = topMatch.item.code;
        matchedName = topMatch.item.name;
        confidence = topMatch.score;
        reason = `Khớp mờ từ vựng đạt ${topMatch.score}%`;
      }

      // 2. Call Gemini AI if low confidence
      if (useAi && aiClient && (!topMatch || topMatch.score < aiThreshold)) {
        try {
          setCurrentAction(`🧠 Gọi AI giải nghĩa dòng ${i + 1}/${maxRows}: "${descValue.substring(0, 20)}..."`);
          
          const candidatesStr = topSuggestions.map(s => `- [Mã ${s.item.code}]: ${s.item.name} (Gợi ý mờ: ${s.score}%)`).join("\n");
          
          const systemPrompt = `Bạn là chuyên gia rà soát danh mục kinh tế xã hội và mã hóa VSIC chuẩn quốc gia.
Nhiệm vụ: Phân loại mô tả hoạt động thực tế thành một mã danh mục phù hợp nhất từ danh sách ứng viên đề xuất hoặc suy luận chuẩn xác.

Danh sách ứng viên từ vựng khớp nhất:
${candidatesStr}

Mô tả hoạt động thực tế cần phân loại: "${descValue}"

Hãy đối chiếu logic, so khớp ngữ nghĩa.
- Nếu một trong các ứng viên khớp ngữ nghĩa xuất sắc (ví dụ 'trồng lúa nước' vs 'trồng lúa', 'nuôi heo' vs 'chăn nuôi lợn'), hãy trả về mã của ứng viên đó.
- Nếu không có ứng viên nào khớp tốt, hãy tự động dùng vốn hiểu biết kinh tế xã hội rộng của bạn để xếp mô tả này vào mã VSIC chuẩn nhất có thể.

Trả về duy nhất chuỗi JSON thuần túy có cấu trúc sau:
{
  "code": "mã danh mục phù hợp nhất (chuỗi)",
  "name": "tên ngành/danh mục chuẩn tương ứng",
  "confidence": "độ tin cậy từ 0 đến 100 (số)",
  "reason": "giải thích ngắn gọn lý do xếp loại trong dưới 20 từ (tiếng Việt)"
}`;

          const response = await aiClient.models.generateContent({
            model: "gemini-3.5-flash",
            contents: `Phân tích và mã hóa: "${descValue}"`,
            config: {
              systemInstruction: systemPrompt,
              responseMimeType: "application/json",
              temperature: 0.1
            }
          });

          const jsonText = (response.text || "").trim();
          const parsed = JSON.parse(jsonText);

          if (parsed && parsed.code) {
            matchedCode = String(parsed.code).trim();
            matchedName = String(parsed.name || vsicRawData[matchedCode] || "Do AI định nghĩa").trim();
            confidence = Math.round(parseFloat(parsed.confidence) || 80);
            method = "Trí tuệ nhân tạo Gemini";
            reason = parsed.reason || "AI tự động liên hợp ngữ nghĩa chuẩn xác";
          }
        } catch (aiErr) {
          console.warn("Gemini Match error at row", i, aiErr);
          // Fallback to best fuzzy match if API fails or rate limits
          reason += " (Lỗi gọi AI: rơi về so khớp mờ)";
        }
      }

      tempResults.push({
        rowIdx: i,
        rowNum: i + 1,
        originalDesc: descValue,
        matchedCode,
        matchedName,
        confidence,
        method,
        reason,
        suggestions: topSuggestions.map(s => ({ code: s.item.code, name: s.item.name, score: s.score }))
      });

      setProgress(Math.round(((i + 1) / maxRows) * 100));
      await new Promise(r => setTimeout(r, 10)); // Yield to prevent lockup
    }

    // Compute stats
    const total = tempResults.length;
    const matched = tempResults.filter(r => r.matchedCode).length;
    const unmatched = total - matched;
    const avgConfidence = total > 0 ? Math.round(tempResults.reduce((sum, r) => sum + r.confidence, 0) / total) : 0;

    setMatchResults(tempResults);
    setCatalogMatchStats({ total, matched, unmatched, avgConfidence });
    setIsMatching(false);
    setShowResults(true);
  };

  // Override suggestion manually
  const handleManualOverride = (rowIdx: number, newCode: string, newName: string) => {
    setMatchResults(prev => prev.map(item => {
      if (item.rowIdx === rowIdx) {
        return {
          ...item,
          matchedCode: newCode,
          matchedName: newName,
          confidence: 100,
          method: "Người dùng điều chỉnh thủ công",
          reason: "Chỉnh sửa trực tiếp từ bảng gợi ý ứng viên"
        };
      }
      return item;
    }));
  };

  // Push matched results into parent's mainData state
  const handleApplyToMainData = () => {
    if (matchResults.length === 0) return;

    const dataCopy = [...mainData];
    let appliedCount = 0;

    matchResults.forEach(item => {
      if (item.matchedCode && item.rowIdx < dataCopy.length) {
        dataCopy[item.rowIdx] = {
          ...dataCopy[item.rowIdx],
          [targetColName]: item.matchedCode,
          [targetNameColName]: item.matchedName,
          "_catalog_match_confidence": item.confidence,
          "_catalog_match_method": item.method
        };
        appliedCount++;
      }
    });

    onUpdateMainData(dataCopy);
    alert(`⚡ Bổ sung cột thành công!\n\nĐã cập nhật ${appliedCount} bản ghi với các cột mới: [${targetColName}] và [${targetNameColName}] vào tệp dữ liệu chính.`);
  };

  // Filter matched results
  const filteredResults = useMemo(() => {
    let list = matchResults;

    // Filter by query
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      list = list.filter(r => {
        return (
          r.originalDesc.toLowerCase().includes(q) ||
          r.matchedCode.toLowerCase().includes(q) ||
          r.matchedName.toLowerCase().includes(q) ||
          r.reason.toLowerCase().includes(q)
        );
      });
    }

    // Filter by quality mode
    if (filterMode === "high") {
      list = list.filter(r => r.confidence >= 80);
    } else if (filterMode === "low") {
      list = list.filter(r => r.confidence > 0 && r.confidence < 80);
    } else if (filterMode === "unmatched") {
      list = list.filter(r => !r.matchedCode);
    }

    return list;
  }, [matchResults, searchQuery, filterMode]);

  return (
    <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-sm font-sans">
      {/* Header */}
      <div>
        <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
          <Brain className="w-5 h-5 text-indigo-600 animate-pulse" /> KHỐI 3: BỘ ĐỐI SÁNH DANH MỤC THÔNG MINH (SMART CATALOG MATCHER)
        </h3>
        <p className="text-xs text-slate-500 mt-0.5">
          Sử dụng cơ chế đối sánh mờ từ vựng kết hợp trí tuệ nhân tạo để xếp loại mô tả thực tế (Mô tả ngành nghề, tên hàng hóa, vị trí địa lý...) sang bất kỳ danh mục chuẩn mong muốn nào.
        </p>
      </div>

      {mainData.length > 0 ? (
        <div className="space-y-6">
          {/* Live 10 Row Preview inside SmartCatalogMatcher */}
          <LiveTenRowPreview
            data={mainData}
            columns={columns}
            onUpdateData={onUpdateMainData}
            highlightedIndices={matchResults.filter(item => !item.matchedCode || item.confidence < 60).map(item => item.rowIdx)}
            highlightLabel="Dòng mờ nhạt/chưa khớp"
            title="BẢNG XEM NHANH & SỬA TRỰC TIẾP 10 DÒNG (ĐỐI SÁNH DANH MỤC)"
          />

          {/* SETUP CONTROL PANEL */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            
            {/* Column & Catalog Selectors */}
            <div className="lg:col-span-2 bg-slate-50 border border-slate-200 p-5 rounded-2xl space-y-4">
              <h4 className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5 pb-2 border-b border-slate-200">
                <Sliders className="w-4 h-4 text-indigo-500" /> THIẾT LẬP CỘT NGUỒN VÀ DANH MỤC ĐỐI SÁNH
              </h4>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-700">1. Cột chứa mô tả cần rà soát/mã hóa:</label>
                  <select
                    value={sourceCol}
                    onChange={(e) => setSourceCol(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs text-slate-800 outline-none focus:ring-1 focus:ring-indigo-500 shadow-sm font-semibold"
                  >
                    {columns.map(col => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                  <p className="text-[10px] text-slate-500">Thường là cột chứa text mô tả hoạt động sản xuất kinh doanh của DTV.</p>
                </div>

                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-700">2. Chọn Danh mục đích để đối sánh:</label>
                  <select
                    value={catalogType}
                    onChange={(e) => setCatalogType(e.target.value as any)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs text-slate-800 outline-none focus:ring-1 focus:ring-indigo-500 shadow-sm font-semibold"
                  >
                    <option value="vsic">Danh mục Ngành Kinh Tế Chuẩn Quốc Gia (VSIC)</option>
                    <option value="custom">Tự dán Danh mục tùy biến (Địa bàn, Sản phẩm, etc.)</option>
                  </select>
                  <p className="text-[10px] text-slate-500">VSIC chuẩn quốc gia gồm {vsicCatalogList.length} ngành 4 &amp; 5 cấp.</p>
                </div>
              </div>

              {/* Custom Catalog Text Area */}
              {catalogType === "custom" && (
                <div className="space-y-1.5 animate-fade-in">
                  <label className="text-xs font-bold text-slate-700 flex items-center gap-1">
                    <Database className="w-3.5 h-3.5 text-indigo-500" /> Dán dữ liệu danh mục tùy biến của bạn (Dạng CSV: Mã, Tên):
                  </label>
                  <textarea
                    rows={4}
                    value={customCatalogText}
                    onChange={(e) => setCustomCatalogText(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl p-3 text-xs font-mono text-slate-700 outline-none focus:ring-1 focus:ring-indigo-500 shadow-inner"
                    placeholder="Mã, Tên danh mục..."
                  />
                  <p className="text-[10px] text-slate-500 font-sans">
                    Hệ thống sẽ tự động phân tách danh mục theo từng dòng. Đã nạp thành công <strong>{customCatalogList.length}</strong> danh mục đích.
                  </p>
                </div>
              )}

              {/* Name output config */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-1">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-700">Đặt tên cột Mã ngành sinh ra:</label>
                  <input
                    type="text"
                    value={targetColName}
                    onChange={(e) => setTargetColName(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-mono text-slate-800"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-700">Đặt tên cột Tên ngành chuẩn sinh ra:</label>
                  <input
                    type="text"
                    value={targetNameColName}
                    onChange={(e) => setTargetNameColName(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-mono text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* AI Engine & Batch Configuration */}
            <div className="bg-slate-50 border border-slate-200 p-5 rounded-2xl flex flex-col justify-between gap-4">
              <div className="space-y-3">
                <h4 className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5 pb-2 border-b border-slate-200">
                  <Brain className="w-4 h-4 text-emerald-500" /> TRÍ TUỆ NHÂN TẠO &amp; BATCH SCAN
                </h4>

                {/* AI Toggle */}
                <div className="bg-white border border-slate-200 p-3 rounded-xl space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700">Kích hoạt AI Gemini:</span>
                    <input
                      type="checkbox"
                      checked={useAi}
                      onChange={(e) => setUseAi(e.target.checked)}
                      className="w-4 h-4 accent-indigo-600 cursor-pointer"
                    />
                  </div>
                  <p className="text-[10px] text-slate-500">
                    Sử dụng mô hình ngôn ngữ lớn để suy luận ngữ nghĩa khi bộ lọc mờ không đạt kết quả tin cậy.
                  </p>
                </div>

                {/* AI Threshold slider */}
                {useAi && (
                  <div className="space-y-1 animate-fade-in">
                    <div className="flex justify-between text-[11px] font-bold">
                      <span className="text-slate-700">Gọi AI khi độ tin cậy mờ dưới:</span>
                      <span className="text-indigo-600 font-mono">{aiThreshold}%</span>
                    </div>
                    <input
                      type="range"
                      min="10"
                      max="100"
                      step="5"
                      value={aiThreshold}
                      onChange={(e) => setAiThreshold(parseInt(e.target.value, 10))}
                      className="w-full accent-indigo-600 h-1.5 bg-slate-200 rounded-lg cursor-pointer"
                    />
                  </div>
                )}

                {/* Limit selection */}
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-700">Giới hạn số bản ghi rà quét:</label>
                  <select
                    value={batchLimit}
                    onChange={(e) => setBatchLimit(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-lg p-1.5 text-xs"
                  >
                    <option value="10">10 dòng đầu tiên (Quét thử nhanh)</option>
                    <option value="50">50 dòng đầu tiên (Khuyên dùng)</option>
                    <option value="200">200 dòng tiếp diễn</option>
                    <option value="all">Chạy toàn bộ tệp (Cần thời gian nếu có gọi AI)</option>
                  </select>
                </div>
              </div>

              {/* Start matching Button */}
              <button
                onClick={handleStartMatch}
                disabled={isMatching}
                className="w-full bg-gradient-to-r from-indigo-500 to-indigo-600 hover:from-indigo-600 hover:to-indigo-700 text-white font-bold text-xs py-3 px-4 rounded-xl cursor-pointer transition-all flex items-center justify-center gap-1.5 border-0 shadow-md active:scale-95 disabled:opacity-50"
              >
                <Zap className="w-4 h-4 text-amber-300" />
                {isMatching ? `Đang đối sánh (${progress}%)...` : "⚡ CHẠY ĐỐI SÁNH DANH MỤC"}
              </button>
            </div>
          </div>

          {/* PROGRESS DISPLAY */}
          {isMatching && (
            <div className="bg-indigo-50 border border-indigo-100 rounded-2xl p-5 space-y-3 animate-pulse">
              <div className="flex justify-between items-center text-xs font-bold text-indigo-800">
                <span>{currentAction}</span>
                <span>{progress}%</span>
              </div>
              <div className="w-full bg-indigo-200 rounded-full h-2.5 overflow-hidden">
                <div className="bg-indigo-600 h-2.5 rounded-full transition-all duration-150" style={{ width: `${progress}%` }}></div>
              </div>
            </div>
          )}

          {/* MATCHING RESULTS INTERACTIVE TABLE */}
          {showResults && (
            <div className="border border-slate-200 rounded-2xl p-5 space-y-4 animate-fade-in bg-white shadow-sm">
              
              {/* Statistical Summary Row */}
              {catalogMatchStats && (
                <div className="grid grid-cols-1 sm:grid-cols-4 gap-4 bg-slate-50 border border-slate-200 p-4 rounded-xl text-center">
                  <div className="space-y-0.5">
                    <span className="text-[10px] uppercase font-bold text-slate-500 block">Tổng số rà quét</span>
                    <strong className="text-lg font-mono text-slate-800">{catalogMatchStats.total}</strong>
                  </div>
                  <div className="space-y-0.5 border-t sm:border-t-0 sm:border-l border-slate-250 pt-2 sm:pt-0">
                    <span className="text-[10px] uppercase font-bold text-slate-500 block">Đã tìm thấy mã</span>
                    <strong className="text-lg font-mono text-emerald-600">{catalogMatchStats.matched}</strong>
                  </div>
                  <div className="space-y-0.5 border-t sm:border-t-0 sm:border-l border-slate-250 pt-2 sm:pt-0">
                    <span className="text-[10px] uppercase font-bold text-slate-500 block">Chưa phân loại</span>
                    <strong className="text-lg font-mono text-red-500">{catalogMatchStats.unmatched}</strong>
                  </div>
                  <div className="space-y-0.5 border-t sm:border-t-0 sm:border-l border-slate-250 pt-2 sm:pt-0">
                    <span className="text-[10px] uppercase font-bold text-slate-500 block">Độ tin cậy trung bình</span>
                    <strong className="text-lg font-mono text-indigo-700">{catalogMatchStats.avgConfidence}%</strong>
                  </div>
                </div>
              )}

              {/* Table Controls */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pt-2">
                <div className="flex flex-wrap gap-2">
                  <button
                    onClick={() => setFilterMode("all")}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                      filterMode === "all" ? "bg-slate-800 text-white" : "bg-slate-100 hover:bg-slate-200 text-slate-600"
                    }`}
                  >
                    Tất cả ({matchResults.length})
                  </button>
                  <button
                    onClick={() => setFilterMode("high")}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                      filterMode === "high" ? "bg-emerald-600 text-white" : "bg-emerald-50 hover:bg-emerald-100 text-emerald-700"
                    }`}
                  >
                    Độ tin cậy cao &gt;= 80% ({matchResults.filter(r => r.confidence >= 80).length})
                  </button>
                  <button
                    onClick={() => setFilterMode("low")}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                      filterMode === "low" ? "bg-amber-600 text-white" : "bg-amber-50 hover:bg-amber-100 text-amber-700"
                    }`}
                  >
                    Nghi ngờ/Cần kiểm tra &lt; 80% ({matchResults.filter(r => r.confidence > 0 && r.confidence < 80).length})
                  </button>
                  <button
                    onClick={() => setFilterMode("unmatched")}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                      filterMode === "unmatched" ? "bg-red-600 text-white" : "bg-red-50 hover:bg-red-100 text-red-700"
                    }`}
                  >
                    Không khớp ({matchResults.filter(r => !r.matchedCode).length})
                  </button>
                </div>

                {/* Filter Search */}
                <div className="relative max-w-xs w-full sm:w-auto">
                  <input
                    type="text"
                    placeholder="Tìm trong kết quả đối sánh..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-300 rounded-lg pl-8 pr-3 py-1.5 text-xs text-slate-800 outline-none focus:bg-white focus:ring-1 focus:ring-indigo-500"
                  />
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                </div>
              </div>

              {/* Table Data list */}
              <div className="overflow-x-auto border border-slate-200 rounded-xl">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-50 text-slate-600 border-b border-slate-200">
                      <th className="py-2.5 px-3 font-bold text-center w-12">STT</th>
                      <th className="py-2.5 px-3 font-bold w-16">Dòng số</th>
                      <th className="py-2.5 px-3 font-bold">Mô tả thực tế của DTV</th>
                      <th className="py-2.5 px-3 font-bold w-28">Mã khớp</th>
                      <th className="py-2.5 px-3 font-bold">Danh mục khớp đề xuất</th>
                      <th className="py-2.5 px-3 font-bold text-center w-16">Tin cậy</th>
                      <th className="py-2.5 px-3 font-bold w-36">Phương thức đối khớp / Chú thích</th>
                      <th className="py-2.5 px-3 font-bold text-center w-40">Điều chỉnh thủ công</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredResults.map((item, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                        <td className="py-2 px-3 text-center text-slate-400 font-mono">{idx + 1}</td>
                        <td className="py-2 px-3 font-bold text-slate-700 font-mono">#{item.rowNum}</td>
                        <td className="py-2 px-3 font-medium text-slate-800 max-w-xs truncate" title={item.originalDesc}>
                          {item.originalDesc}
                        </td>
                        <td className="py-2 px-3 font-mono font-extrabold text-indigo-700">
                          {item.matchedCode || (
                            <span className="text-red-500 text-[10px] font-bold font-sans flex items-center gap-0.5">
                              <AlertTriangle className="w-3 h-3" /> TRỐNG
                            </span>
                          )}
                        </td>
                        <td className="py-2 px-3 text-slate-800 truncate max-w-xs" title={item.matchedName}>
                          {item.matchedName}
                        </td>
                        <td className="py-2 px-3 text-center">
                          <span className={`px-1.5 py-0.5 rounded font-mono font-bold text-[10px] ${
                            item.confidence >= 80 
                              ? "bg-emerald-100 text-emerald-800 border border-emerald-200" 
                              : item.confidence > 0 
                                ? "bg-amber-100 text-amber-800 border border-amber-200"
                                : "bg-red-100 text-red-800 border border-red-200"
                          }`}>
                            {item.confidence}%
                          </span>
                        </td>
                        <td className="py-2 px-3 text-[10.5px] text-slate-500 leading-tight">
                          <strong className="block text-[10px] text-slate-600 font-mono">{item.method}</strong>
                          <span className="italic">{item.reason}</span>
                        </td>
                        <td className="py-2 px-3">
                          {item.suggestions && item.suggestions.length > 1 ? (
                            <select
                              onChange={(e) => {
                                const selected = item.suggestions.find((s: any) => s.code === e.target.value);
                                if (selected) {
                                  handleManualOverride(item.rowIdx, selected.code, selected.name);
                                }
                              }}
                              className="w-full bg-slate-50 hover:bg-white border border-slate-300 rounded px-1.5 py-1 text-[10.5px] text-slate-700 outline-none cursor-pointer"
                              defaultValue={item.matchedCode}
                            >
                              <option value={item.matchedCode}>-- Chọn từ gợi ý khác --</option>
                              {item.suggestions.map((s: any, sIdx: number) => (
                                <option key={sIdx} value={s.code}>
                                  [{s.code}] {s.name.substring(0, 20)}... ({s.score}%)
                                </option>
                              ))}
                            </select>
                          ) : (
                            <span className="text-[10px] text-slate-400 italic block text-center">Không có ứng viên phụ</span>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* ACTION COMMAND BAR */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pt-3 border-t border-slate-100">
                <div className="text-xs text-slate-500">
                  * Nhấp nút bổ sung cột để chèn kết quả trực tiếp vào tệp dữ liệu chính của bạn để lưu trữ/sử dụng tiếp.
                </div>
                <div className="flex gap-3">
                  <button
                    onClick={() => {
                      // Prepare data for Excel
                      const excelRows = matchResults.map(r => ({
                        "Dòng số": r.rowNum,
                        "Mô tả gốc": r.originalDesc,
                        "Mã ngành đối sánh": r.matchedCode,
                        "Tên ngành đối sánh chuẩn": r.matchedName,
                        "Hệ số tin cậy %": r.confidence,
                        "Phương pháp đối khớp": r.method,
                        "Giải thích chi tiết": r.reason
                      }));
                      if (onExportExcel) {
                        onExportExcel(excelRows, `Doi_Sanh_Danh_Muc_Thong_Minh_${fileNameClean()}`);
                      } else {
                        const worksheet = XLSX.utils.json_to_sheet(excelRows);
                        const workbook = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(workbook, worksheet, "Đối sánh danh mục");
                        XLSX.writeFile(workbook, `Báo_Cáo_Đối_Sánh_Danh_Mục_${fileNameClean()}.xlsx`);
                      }
                    }}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs py-2 px-4 rounded-xl cursor-pointer border border-slate-200 flex items-center gap-1 transition-all"
                  >
                    <Download className="w-3.5 h-3.5" /> Xuất tệp báo cáo (.xlsx)
                  </button>
                  <button
                    onClick={handleApplyToMainData}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs py-2 px-5 rounded-xl cursor-pointer border-0 flex items-center gap-1.5 transition-all shadow-md active:scale-95"
                  >
                    <ArrowRightLeft className="w-3.5 h-3.5" /> ⚡ BỔ SUNG CỘT ĐỐI SÁNH VÀO DỮ LIỆU CHÍNH
                  </button>
                </div>
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

function fileNameClean() {
  return new Date().toISOString().slice(0, 10).replace(/-/g, "");
}
