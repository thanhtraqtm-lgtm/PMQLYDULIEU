import React, { useState, useEffect, useMemo } from "react";
import { 
  CheckSquare, Sliders, Sparkles, Download, Upload, 
  Trash2, AlertTriangle, Play, Save, Activity, Mic, MicOff
} from "lucide-react";
import { GoogleGenAI } from "@google/genai";
import { MainDataInlinePreview } from "./MainDataInlinePreview";
import { getFlexibleValue, normalizeAiExpression } from "../utils/sharedHelpers";

// Local types
interface LogicRule {
  col: string;
  op: string;
  val: string;
  isFieldCompare?: boolean;
}

interface LogicCheckingProps {
  mainData: any[];
  columns: string[];
  setMainData: (data: any[]) => void;
  setColumns: (cols: string[]) => void;
  fileName: string;
  setFileName: (name: string) => void;
  mapping: any;
  onExportExcel: () => void;
  saveAppState: (newState: any, isCritical?: boolean) => void;
  setActiveTab: (tab: string) => void;
  setLoading: (loading: boolean) => void;
  setProgress: (progress: number) => void;
  setStatusMessage: (msg: string) => void;
  rawImportedData: any[];
  customColConfigs: any;
  quickReportManganhCol: string;
  setQuickReportManganhCol: (col: string) => void;
  quickReportXaCol: string;
  setQuickReportXaCol: (col: string) => void;
  quickReportDoanhThuCol: string;
  setQuickReportDoanhThuCol: (col: string) => void;
  quickReportLaoDongCol: string;
  setQuickReportLaoDongCol: (col: string) => void;
  handleQuickReport: (level: number, optManganh?: string, optXa?: string, optDoanhThu?: string, optLaoDong?: string) => Promise<any>;
}

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

const chunkProcess = async <T, R>(
  array: T[],
  size: number,
  processFn: (item: T, index: number) => R,
  onProgress?: (percent: number) => void
): Promise<R[]> => {
  const result: R[] = [];
  const len = array.length;
  if (len === 0) return [];
  
  for (let i = 0; i < len; i += size) {
    const chunk = array.slice(i, i + size);
    for (let j = 0; j < chunk.length; j++) {
      result.push(processFn(chunk[j], i + j));
    }
    if (onProgress) {
      onProgress(Math.min(100, Math.round((i / len) * 100)));
    }
    await new Promise<void>(resolve => setTimeout(resolve, 0));
  }
  return result;
};

export default function LogicChecking({
  mainData,
  columns,
  setMainData,
  setColumns,
  fileName,
  setFileName,
  mapping,
  onExportExcel,
  saveAppState,
  setActiveTab,
  setLoading,
  setProgress,
  setStatusMessage,
  rawImportedData,
  customColConfigs,
  quickReportManganhCol,
  setQuickReportManganhCol,
  quickReportXaCol,
  setQuickReportXaCol,
  quickReportDoanhThuCol,
  setQuickReportDoanhThuCol,
  quickReportLaoDongCol,
  setQuickReportLaoDongCol,
  handleQuickReport
}: LogicCheckingProps) {

  // Quy tắc kiểm tra logic đa điều kiện
  const [ifRules, setIfRules] = useState<LogicRule[]>([]);
  const [thenRules, setThenRules] = useState<LogicRule[]>([]);
  const [ifCombine, setIfCombine] = useState<"AND" | "OR">("AND");
  const [thenCombine, setThenCombine] = useState<"AND" | "OR">("AND");
  const [logicRuleMode, setLogicRuleMode] = useState<"conflict" | "must_satisfy">("conflict");
  const [logicFilterMode, setLogicFilterMode] = useState<"all" | "if_satisfied" | "violated">("if_satisfied");

  // Trí tuệ Nhân tạo - Học và lưu lệnh qua AI
  const [aiRulePrompt, setAiRulePrompt] = useState<string>("");
  const [isRecordingRuleMic, setIsRecordingRuleMic] = useState(false);

  const toggleMicRule = () => {
    const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
    if (!SpeechRecognition) {
      alert("Trình duyệt không hỗ trợ nhận diện giọng nói (Web Speech API). Hãy dùng Google Chrome hoặc Microsoft Edge.");
      return;
    }

    if (isRecordingRuleMic) {
      setIsRecordingRuleMic(false);
      return;
    }

    const rec = new SpeechRecognition();
    rec.lang = "vi-VN";
    rec.interimResults = false;
    rec.maxAlternatives = 1;

    rec.onstart = () => {
      setIsRecordingRuleMic(true);
    };

    rec.onerror = (e: any) => {
      console.error(e);
      setIsRecordingRuleMic(false);
    };

    rec.onend = () => {
      setIsRecordingRuleMic(false);
    };

    rec.onresult = (event: any) => {
      const transcript = event.results[0][0].transcript;
      if (transcript) {
        setAiRulePrompt(prev => {
          const trimmed = prev.trim();
          return trimmed ? `${trimmed} ${transcript}` : transcript;
        });
      }
    };

    rec.start();
  };
  const [aiTranslatedExpression, setAiTranslatedExpression] = useState<string>("");
  const [customRuleName, setCustomRuleName] = useState<string>("");
  const [aiScanMetrics, setAiScanMetrics] = useState<{
    total: number;
    violated: number;
    passed: number;
    violatedPercent: string;
    passedPercent: string;
    expression: string;
    prompt: string;
  } | null>(null);

  const [savedAiRules, setSavedAiRules] = useState<{ id: string; name: string; prompt: string; expression: string }[]>(() => {
    try {
      const saved = localStorage.getItem("vsic_saved_ai_rules");
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      return [];
    }
  });

  // Quy tắc mới cho Logic
  const [newIfRule, setNewIfRule] = useState<LogicRule>({ col: "", op: "==", val: "", isFieldCompare: false });
  const [newThenRule, setNewThenRule] = useState<LogicRule>({ col: "", op: "==", val: "", isFieldCompare: false });

  // Bộ lọc dữ liệu logic
  const filteredLogicData = useMemo(() => {
    if (!mainData || mainData.length === 0) return [];
    const hasBeenScanned = mainData.some(row => "_satisfiesIf" in row || "_violated" in row);
    if (!hasBeenScanned) return [];

    if (logicFilterMode === "if_satisfied") {
      return mainData.filter(row => row._satisfiesIf === true);
    }
    if (logicFilterMode === "violated") {
      return mainData.filter(row => row._violated === true);
    }
    return mainData.filter(row => row._satisfiesIf === true || row._violated === true);
  }, [mainData, logicFilterMode]);

  const handleLogicRuleAdd = (type: "if" | "then") => {
    if (type === "if") {
      if (!newIfRule.col) {
        alert("Vui lòng chọn cột điều kiện NẾU!");
        return;
      }
      setIfRules([...ifRules, newIfRule]);
      setNewIfRule({ col: "", op: "==", val: "", isFieldCompare: false });
    } else {
      if (!newThenRule.col) {
        alert("Vui lòng chọn cột điều kiện THÌ PHẢI!");
        return;
      }
      setThenRules([...thenRules, newThenRule]);
      setNewThenRule({ col: "", op: "==", val: "", isFieldCompare: false });
    }
  };

  const handleLogicCheck = async () => {
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi kiểm tra logic.");
      return;
    }
    if (ifRules.length === 0) {
      alert("Hãy định cấu hình ít nhất 1 quy tắc rà soát 'NẾU' ở Bước 1!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Bắt đầu kiểm tra logic đa điều kiện...");
    await sleep(200);

    const checkValue = (rowVal: any, op: string, compareVal: string) => {
      const v1 = String(rowVal !== undefined && rowVal !== null ? rowVal : "").trim();
      const v2 = String(compareVal).trim();

      const v1LC = v1.toLowerCase();
      const v2LC = v2.toLowerCase();

      if (op === "trống") return v1 === "";
      if (op === "không trống") return v1 !== "";

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

    const results = await chunkProcess(
      mainData,
      10000,
      (row, index) => {
        if (!row || typeof row !== 'object') return row;
        const ifMatches = ifRules.map(r => {
          const compVal = r.isFieldCompare ? String(row[r.val] !== undefined && row[r.val] !== null ? row[r.val] : "") : r.val;
          return checkValue(row[r.col], r.op, compVal);
        });
        const satisfiesIf = ifCombine === "AND" 
          ? ifMatches.every(v => v === true) 
          : ifMatches.some(v => v === true);

        let biViPham = false;
        let noteLoi = "";

        const getRuleDescription = (r: LogicRule) => {
          const rightSide = r.isFieldCompare ? `Cột [${r.val}]` : `'${r.val}'`;
          return `(${r.col} ${r.op} ${rightSide})`;
        };

        if (thenRules.length === 0) {
          if (satisfiesIf) {
            biViPham = true;
            const descriptIf = ifRules.map(getRuleDescription).join(` ${ifCombine} `);
            noteLoi = `[ĐÃ TÌM THẤY] Thỏa mãn điều kiện lọc kiểm tra: { ${descriptIf} }; `;
          }
        } else {
          const thenMatches = thenRules.map(r => {
            const compVal = r.isFieldCompare ? String(row[r.val] !== undefined && row[r.val] !== null ? row[r.val] : "") : r.val;
            return checkValue(row[r.col], r.op, compVal);
          });
          const satisfiesThen = thenCombine === "AND"
            ? thenMatches.every(v => v === true)
            : thenMatches.some(v => v === true);

          if (logicRuleMode === "conflict") {
            if (satisfiesIf && satisfiesThen) {
              biViPham = true;
              const descriptIf = ifRules.map(getRuleDescription).join(` ${ifCombine} `);
              const descriptThen = thenRules.map(getRuleDescription).join(` ${thenCombine} `);
              noteLoi = `[MÂU THUẪN LOGIC] Thỏa mãn đồng thời: { NẾU: ${descriptIf} } và { CÓ THÊM: ${descriptThen} }; `;
            }
          } else {
            if (satisfiesIf && !satisfiesThen) {
              biViPham = true;
              const descriptIf = ifRules.map(getRuleDescription).join(` ${ifCombine} `);
              const descriptThen = thenRules.map(getRuleDescription).join(` ${thenCombine} `);
              noteLoi = `[VI PHẠM LOGIC] NẾU thỏa mãn: { ${descriptIf} } THÌ BẮT BUỘC PHẢI THỎA MÃN: { ${descriptThen} }; `;
            }
          }
        }

        return {
          ...row,
          "Loi_Logic": biViPham ? noteLoi : "✅ Đạt",
          "_satisfiesIf": satisfiesIf,
          "_violated": biViPham
        };
      },
      pct => {
        setProgress(pct);
        setStatusMessage(`Đang chạy kiểm tra logic đa điều kiện: ${pct}%...`);
      }
    );

    setMainData(results);
    setColumns(Object.keys(results[0] || {}));
    if (!fileName.startsWith("KiemTraLogic_")) {
      setFileName(`KiemTraLogic_${fileName}`);
    }

    const totalRows = results.length;
    const violatedCount = results.filter((r: any) => r && r._violated === true).length;
    const passedCount = totalRows - violatedCount;

    setAiScanMetrics({
      total: totalRows,
      violated: violatedCount,
      passed: passedCount,
      violatedPercent: totalRows > 0 ? ((violatedCount / totalRows) * 100).toFixed(2) + "%" : "0%",
      passedPercent: totalRows > 0 ? ((passedCount / totalRows) * 100).toFixed(2) + "%" : "0%",
      expression: "Quy tắc rà soát logic đa điều kiện",
      prompt: "Kiểm tra logic thủ công"
    });

    if (violatedCount > 0) {
      setLogicFilterMode("violated");
    } else {
      setLogicFilterMode("all");
    }

    setProgress(100);
    setStatusMessage(`Kiểm tra hoàn tất! Đã phân tích kiểm tra và phát hiện các dòng lỗi.`);
    await sleep(400);
    setLoading(false);
  };

  const handleAiLogicScan = async (overridePrompt?: string) => {
    const activePrompt = overridePrompt || aiRulePrompt;
    if (!activePrompt.trim()) {
      alert("Vui lòng nhập khẩu lệnh rà quét bằng tiếng Việt!");
      return;
    }
    if (mainData.length === 0) {
      alert("Vui lòng nạp dữ liệu chính trước khi thực hiện quét!");
      return;
    }

    const activePromptLC = activePrompt.toLowerCase();
    const isReportRequest = (
      activePromptLC.includes("tổng hợp") || 
      activePromptLC.includes("tong hop") || 
      activePromptLC.includes("báo cáo") || 
      activePromptLC.includes("bao cao") ||
      activePromptLC.includes("phân tích doanh thu")
    ) && (
      activePromptLC.includes("xã") || 
      activePromptLC.includes("xa") || 
      activePromptLC.includes("ngành") || 
      activePromptLC.includes("nganh") ||
      activePromptLC.includes("cơ cấu") ||
      activePromptLC.includes("tỉ trọng") ||
      activePromptLC.includes("tỷ trọng")
    );

    if (isReportRequest) {
      const colManganh = quickReportManganhCol || mapping.manganh || columns.find(c => /mã\s*ngành|manganh|vsic|mã\s*nghe|manghe|ngành/i.test(c)) || "";
      const colXa = quickReportXaCol || mapping.xa || columns.find(c => /xã|xa|địa\s*bàn|dia\s*ban|phường|phuong/i.test(c)) || "";
      const colDoanhThu = quickReportDoanhThuCol || mapping.doanhthu || columns.find(c => /doanh\s*thu|doanhthu|thu\s*nhập|thunhap|tiền|tien/i.test(c)) || "";
      const colLaoDong = quickReportLaoDongCol || mapping.laodong || columns.find(c => /lao\s*động|laodong|người|nguoi|nhân\s*sự|nhansu/i.test(c)) || "";

      if (colManganh) setQuickReportManganhCol(colManganh);
      if (colXa) setQuickReportXaCol(colXa);
      if (colDoanhThu) setQuickReportDoanhThuCol(colDoanhThu);
      if (colLaoDong) setQuickReportLaoDongCol(colLaoDong);

      setLoading(true);
      setProgress(40);
      setStatusMessage("Hệ thống phát hiện lệnh Tổng Hợp Báo Cáo! Đang tự động chuyển sang Tab 'Tổng Hợp Báo Cáo'...");
      
      await sleep(1000);
      setActiveTab("tonghop");
      
      const targetLevel = activePromptLC.includes("cấp 1") || activePromptLC.includes("cap 1") ? 1 : 2;
      setProgress(75);
      setStatusMessage(`Đang chạy hạch toán tổng hợp: Ngành Cấp ${targetLevel} và Địa bàn Xã / Phường...`);
      await sleep(600);

      try {
        await handleQuickReport(targetLevel, colManganh, colXa, colDoanhThu, colLaoDong);
      } catch (err: any) {
        alert("Lỗi rẽ hướng tổng hợp liên kết: " + err.message);
      }
      setLoading(false);
      return;
    }

    const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || "";
    if (!apiKey) {
      alert("Cần cấu hình khoá API VITE_GEMINI_API_KEY trong phần Cài đặt > Khóa bí mật (Secrets) của AI Studio hoặc trong file .env!");
      return;
    }

    setLoading(true);
    setProgress(10);
    setStatusMessage("Trí tuệ nhân tạo đang phân tích và dịch khẩu lệnh...");

    try {
      let expression = "";

      if (overridePrompt) {
        const matched = savedAiRules.find(r => r.prompt === overridePrompt || r.name === overridePrompt);
        if (matched) {
          expression = matched.expression;
        }
      }

      if (!expression) {
        const ai = new GoogleGenAI({
          apiKey: apiKey,
          httpOptions: {
            headers: {
              "User-Agent": "aistudio-build"
            }
          }
        });

        const systemPrompt = `Bạn là chuyên gia chuyển dịch khẩu lệnh tiếng Việt thành biểu thức điều kiện JavaScript chính xác cho bảng dữ liệu.
Mục tiêu là: Dịch yêu cầu tìm kiếm lỗi dữ liệu của người dùng thành một biểu thức logic JavaScript trả về true khi dòng đó bị lỗi hoặc vi phạm điều kiện.
Bạn PHẢI sử dụng biến đối tượng là 'row' để truy cập các cột của dòng.

Các cột dữ liệu hiện tại trong file của người dùng gồm: [${columns.filter(c => !c.startsWith("_")).map(c => `'${c}'`).join(", ")}].
Hãy phân tích ngôn từ của người dùng và khớp chính xác các cột trên. Nếu cột có tiếng Việt, hãy truy cập theo dạng row['Tên Cột'].
Luôn chú ý kiểu dữ liệu (nếu so sánh số, hãy dùng parseFloat(row['Tên Cột']) hoặc so sánh trực tiếp, loại bỏ dấu phẩy ngăn cách hàng nghìn nếu cần).

Ví dụ:
1. "DonGia < 0" -> parseFloat(String(row['DonGia'] || '0').replace(/,/g, '')) < 0
2. "Nhà thuê mượn = 1 nhưng tài sản lớn hơn 0" -> row['Thuê mượn'] == 1 && parseFloat(String(row['Tài sản'] || '0').replace(/,/g, '')) > 0
3. "Số điện thoại bị trống" -> !row['Số điện thoại'] || String(row['Số điện thoại']).trim() === ''

Quy tắc cực kỳ quan trọng:
CHỈ TRẢ VỀ DUY NHẤT CHUỖI BIỂU THỨC LOGIC TRONG JAVASCRIPT ĐỂ ĐƯA VÀO HÀM EVAL/NEW FUNCTION.
KHÔNG giải thích, KHÔNG bọc trong khối mã markdown (\`\`\`), KHÔNG chứa bất cứ từ ngữ thừa thãi hay ký tự nào ngoài biểu thức.`;

        const response = await ai.models.generateContent({
          model: "gemini-2.5-flash",
          contents: `Khẩu lệnh người dùng: "${activePrompt}"\nHãy chuyển dịch thành biểu thức Javascript viết dạng row['Cột']...`,
          config: {
            systemInstruction: systemPrompt,
            temperature: 0.1
          }
        });

        const rawResult = response.text || "";
        expression = rawResult
          .replace(/```javascript/gi, "")
          .replace(/```js/gi, "")
          .replace(/```/g, "")
          .trim();

        if (!expression) {
          throw new Error("Mô hình AI phản hồi trống hoặc không hợp lệ.");
        }
      }

      setAiTranslatedExpression(expression);
      setCustomRuleName(overridePrompt ? "" : `Quy tắc: ${activePrompt.substring(0, 25)}`);
      setStatusMessage("Dịch thuật thành công! Bắt đầu rà quét dữ liệu bằng bộ quét hiệu năng...");
      setProgress(30);
      await sleep(200);

      let violatedCount = 0;
      let passedCount = 0;
      const normalizedExpr = normalizeAiExpression(expression);

      const results = await chunkProcess(
        mainData,
        10000,
        (row, index) => {
          if (!row || typeof row !== 'object') return row;
          
          let biViPham = false;
          try {
            const evaluator = new Function("row", "getFlexibleValue", `
              try {
                return !!(${normalizedExpr});
              } catch(e) {
                return false;
              }
            `);
            biViPham = evaluator(row, getFlexibleValue);
          } catch (err) {
            biViPham = false;
          }

          if (biViPham) {
            violatedCount++;
          } else {
            passedCount++;
          }

          return {
            ...row,
            "Loi_Logic": biViPham ? `[LỖI AI-LỆNH]: thỏa mãn quy tắc "${activePrompt}"` : "✅ Đạt",
            "_satisfiesIf": true,
            "_violated": biViPham
          };
        },
        pct => {
          setProgress(Math.round(30 + (pct * 0.7)));
          setStatusMessage(`Trí tuệ nhân tạo đang quét dữ liệu: ${pct}%...`);
        }
      );

      setMainData(results);
      setColumns(Object.keys(results[0] || {}));
      if (!fileName.startsWith("QuetAI_")) {
        setFileName(`QuetAI_${fileName}`);
      }

      const totalRows = results.length;
      setAiScanMetrics({
        total: totalRows,
        violated: violatedCount,
        passed: passedCount,
        violatedPercent: totalRows > 0 ? ((violatedCount / totalRows) * 100).toFixed(2) + "%" : "0%",
        passedPercent: totalRows > 0 ? ((passedCount / totalRows) * 100).toFixed(2) + "%" : "0%",
        expression: expression,
        prompt: activePrompt
      });

      if (violatedCount > 0) {
        setLogicFilterMode("violated");
      } else {
        setLogicFilterMode("all");
      }

      setProgress(100);
      setStatusMessage(`Đã rà quét hoàn tất bằng AI dựa trên biểu thức: "${expression}"`);
      await sleep(400);
      setLoading(false);

      setTimeout(() => {
        const docEl = document.getElementById("ai-scan-summary-section");
        if (docEl) {
          docEl.scrollIntoView({ behavior: "smooth", block: "center" });
        }
      }, 250);

    } catch (err: any) {
      console.error(err);
      alert(`Đã xảy ra lỗi khi quét bằng AI: ${err.message || err}`);
      setLoading(false);
    }
  };

  const handleSaveAiRule = () => {
    if (!aiRulePrompt.trim() || !aiTranslatedExpression.trim()) {
      alert("Chưa có biểu thức nào được AI dịch thành công để lưu cả!");
      return;
    }
    const ruleName = customRuleName.trim() || `Luật rà quét ${aiRulePrompt.substring(0, 20)}...`;
    
    const newRule = {
      id: String(Date.now()),
      name: ruleName,
      prompt: aiRulePrompt.trim(),
      expression: aiTranslatedExpression.trim()
    };

    const updated = [newRule, ...savedAiRules];
    setSavedAiRules(updated);
    localStorage.setItem("vsic_saved_ai_rules", JSON.stringify(updated));
    alert(`Đã lưu thành công quy tắc "${ruleName}" vào bộ nhớ nhanh của trình duyệt!`);
  };

  const handleExportAiRules = () => {
    if (savedAiRules.length === 0) {
      alert("Chưa có danh sách quy tắc học lệnh nào được lưu để xuất!");
      return;
    }
    try {
      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(savedAiRules, null, 2));
      const downloadAnchor = document.createElement("a");
      downloadAnchor.setAttribute("href", dataStr);
      downloadAnchor.setAttribute("download", "AI_Logic_Rules_Backup.json");
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      downloadAnchor.remove();
    } catch (err: any) {
      alert(`Lỗi xuất file: ${err.message || err}`);
    }
  };

  const handleImportAiRules = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileReader = new FileReader();
    if (e.target.files && e.target.files[0]) {
      fileReader.readAsText(e.target.files[0], "UTF-8");
      fileReader.onload = (event) => {
        try {
          const parsed = JSON.parse(event.target?.result as string);
          if (Array.isArray(parsed)) {
            const valid = parsed.every(item => item.id && item.name && item.prompt && item.expression);
            if (!valid) {
              alert("Lỗi: Các quy tắc trong file chứa định dạng không hợp lệ!");
              return;
            }
            const merged = [...parsed, ...savedAiRules.filter(existing => !parsed.some(p => p.id === existing.id))];
            setSavedAiRules(merged);
            localStorage.setItem("vsic_saved_ai_rules", JSON.stringify(merged));
            alert(`Đã nạp và đồng bộ thành công ${parsed.length} quy tắc học lệnh thông minh bằng AI!`);
          } else {
            alert("Tệp tin JSON tải lên không hợp lệ (phải là một danh sách các quy tắc)!");
          }
        } catch (err: any) {
          alert(`Nạp lỗi! Đọc nội dung bị lỗi: ${err.message || err}`);
        }
      };
    }
  };

  return (
    <div className="space-y-6 animate-fade-in">
      <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-sm">
        <div>
          <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
            <CheckSquare className="w-5 h-5 text-emerald-600" /> TRÌNH THIẾT LẬP QUY TẮC KIỂM TRA DỮ LIỆU
          </h3>
          <p className="text-xs text-slate-500">Thiết lập quy tắc kiểm tra thông minh theo logic: Nếu (Điều kiện 1) xảy ra, thì (Điều kiện 2) bắt buộc phải đúng.</p>
        </div>

        {mainData.length > 0 ? (
          <div className="space-y-6 border-t border-slate-200 pt-6">

            {/* KHU VỰC THIẾT LẬP CÔNG THỨC LOGIC THỦ CÔNG */}
            <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-sm">
              <div className="border-b border-slate-100 pb-3">
                <h4 className="text-sm font-extrabold text-slate-800 flex items-center gap-2">
                  <Sliders className="w-5 h-5 text-emerald-600" /> BƯỚC 1: XÂY DỰNG QUY TẮC LOGIC THỦ CÔNG
                </h4>
                <p className="text-xs text-slate-500 mt-1">
                  Tạo các quy tắc "NẾU ... THÌ BẮT BUỘC PHẢI ..." để rà quét toàn bộ file gốc. Các toán tử tự động hỗ trợ so khớp số và chữ.
                </p>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* CỘT TRÁI: ĐIỀU KIỆN TIỀN ĐỀ (NẾU) */}
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4.5 space-y-4">
                  <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                    <span className="text-xs font-bold text-slate-800 uppercase tracking-wider font-mono flex items-center gap-1">
                      🔴 VẾ "NẾU" (ĐIỀU KIỆN TIỀN ĐỀ)
                    </span>
                    <div className="flex items-center gap-1.5 text-xs">
                      <span className="text-slate-500 text-[11px]">Kết hợp:</span>
                      <select
                        value={ifCombine}
                        onChange={(e) => setIfCombine(e.target.value as "AND" | "OR")}
                        className="bg-white border border-slate-300 rounded px-1.5 py-0.5 text-xs font-bold text-indigo-700 outline-none cursor-pointer"
                      >
                        <option value="AND">TẤT CẢ (AND)</option>
                        <option value="OR">MỘT TRONG (OR)</option>
                      </select>
                    </div>
                  </div>

                  {/* Bộ thiết lập quy tắc NẾU mới */}
                  <div className="space-y-3 bg-white p-3 rounded-lg border border-slate-200">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                      <div className="space-y-1">
                        <label className="text-[10px] uppercase font-bold text-slate-500">Chọn cột nguồn:</label>
                        <select
                          value={newIfRule.col}
                          onChange={(e) => setNewIfRule({ ...newIfRule, col: e.target.value })}
                          className="w-full bg-slate-50 border border-slate-300 rounded px-2 py-1 text-xs text-slate-800 font-medium"
                        >
                          <option value="">-- Chọn cột --</option>
                          {columns.map(col => (
                            <option key={col} value={col}>{col}</option>
                          ))}
                        </select>
                      </div>

                      <div className="space-y-1">
                        <label className="text-[10px] uppercase font-bold text-slate-500">Phép so sánh:</label>
                        <select
                          value={newIfRule.op}
                          onChange={(e) => setNewIfRule({ ...newIfRule, op: e.target.value })}
                          className="w-full bg-slate-50 border border-slate-300 rounded px-2 py-1 text-xs text-slate-800 font-medium"
                        >
                          <option value="==">Bằng (==)</option>
                          <option value="!=">Khác (!=)</option>
                          <option value=">">Lớn hơn (&gt;)</option>
                          <option value="<">Nhỏ hơn (&lt;)</option>
                          <option value=">=">Lớn hơn hoặc bằng (&gt;=)</option>
                          <option value="<=">Nhỏ hơn hoặc bằng (&lt;=)</option>
                          <option value="chứa">Chứa từ (chứa)</option>
                          <option value="không chứa">Không chứa từ</option>
                          <option value="trống">Để trống (rỗng)</option>
                          <option value="không trống">Không để trống</option>
                        </select>
                      </div>
                    </div>

                    {newIfRule.op !== "trống" && newIfRule.op !== "không trống" && (
                      <div className="space-y-2">
                        <div className="flex items-center gap-1.5">
                          <input
                            type="checkbox"
                            id="newIfRule_isFieldCompare"
                            checked={newIfRule.isFieldCompare}
                            onChange={(e) => setNewIfRule({ ...newIfRule, isFieldCompare: e.target.checked, val: "" })}
                            className="cursor-pointer"
                          />
                          <label htmlFor="newIfRule_isFieldCompare" className="text-[11px] text-slate-600 font-bold cursor-pointer">So sánh với giá trị cột khác</label>
                        </div>

                        <div className="space-y-1">
                          <label className="text-[10px] uppercase font-bold text-slate-500">Giá trị so sánh:</label>
                          {newIfRule.isFieldCompare ? (
                            <select
                              value={newIfRule.val}
                              onChange={(e) => setNewIfRule({ ...newIfRule, val: e.target.value })}
                              className="w-full bg-slate-50 border border-slate-300 rounded px-2 py-1 text-xs text-slate-800 font-medium"
                            >
                              <option value="">-- Chọn cột để so sánh --</option>
                              {columns.map(col => (
                                <option key={col} value={col}>{col}</option>
                              ))}
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={newIfRule.val}
                              onChange={(e) => setNewIfRule({ ...newIfRule, val: e.target.value })}
                              placeholder="Nhập giá trị chữ hoặc số..."
                              className="w-full bg-slate-50 border border-slate-300 rounded px-2 py-1 text-xs text-slate-800"
                            />
                          )}
                        </div>
                      </div>
                    )}

                    <button
                      onClick={() => handleLogicRuleAdd("if")}
                      className="w-full bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-xs py-1.5 rounded border border-indigo-200 cursor-pointer transition-all flex items-center justify-center gap-1"
                    >
                      ➕ Thêm điều kiện NẾU
                    </button>
                  </div>

                  {/* Danh sách quy tắc NẾU đã thêm */}
                  <div className="space-y-1.5">
                    <span className="text-[10.5px] font-bold text-slate-500 block uppercase">Điều kiện đang áp dụng:</span>
                    {ifRules.length === 0 ? (
                      <div className="text-[11px] text-slate-400 italic bg-white border border-slate-150 p-3 rounded text-center">
                        Chưa thiết lập điều kiện nào. Tất cả dòng sẽ được rà soát.
                      </div>
                    ) : (
                      <div className="space-y-1 max-h-[140px] overflow-y-auto">
                        {ifRules.map((rule, idx) => (
                          <div key={idx} className="bg-white border border-slate-200 px-2.5 py-1.5 rounded flex items-center justify-between text-xs font-mono">
                            <span className="text-slate-700">
                              [{rule.col}] <strong className="text-indigo-600">{rule.op}</strong> {rule.isFieldCompare ? `Cột [${rule.val}]` : `'${rule.val || "rỗng"}'`}
                            </span>
                            <button
                              onClick={() => setIfRules(ifRules.filter((_, i) => i !== idx))}
                              className="text-slate-400 hover:text-red-500 p-0.5 rounded hover:bg-slate-100 transition-colors cursor-pointer"
                            >
                              ✕
                            </button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>

                {/* CỘT PHẢI: ĐIỀU KIỆN RÀNG BUỘC (THÌ BẮT BUỘC PHẢI THỎA MÃN) */}
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4.5 space-y-4">
                  <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                    <span className="text-xs font-bold text-slate-800 uppercase tracking-wider font-mono flex items-center gap-1">
                      🟢 VẾ "THÌ PHẢI" (RÀNG BUỘC BẮT BUỘC)
                    </span>
                    <div className="flex items-center gap-1.5 text-xs">
                      <span className="text-slate-500 text-[11px]">Kết hợp:</span>
                      <select
                        value={thenCombine}
                        onChange={(e) => setThenCombine(e.target.value as "AND" | "OR")}
                        className="bg-white border border-slate-300 rounded px-1.5 py-0.5 text-xs font-bold text-indigo-700 outline-none cursor-pointer"
                      >
                        <option value="AND">TẤT CẢ (AND)</option>
                        <option value="OR">MỘT TRONG (OR)</option>
                      </select>
                    </div>
                  </div>

                  {/* Bộ thiết lập quy tắc THÌ PHẢI mới */}
                  <div className="space-y-3 bg-white p-3 rounded-lg border border-slate-200">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                      <div className="space-y-1">
                        <label className="text-[10px] uppercase font-bold text-slate-500">Chọn cột đích:</label>
                        <select
                          value={newThenRule.col}
                          onChange={(e) => setNewThenRule({ ...newThenRule, col: e.target.value })}
                          className="w-full bg-slate-50 border border-slate-300 rounded px-2 py-1 text-xs text-slate-800 font-medium"
                        >
                          <option value="">-- Chọn cột --</option>
                          {columns.map(col => (
                            <option key={col} value={col}>{col}</option>
                          ))}
                        </select>
                      </div>

                      <div className="space-y-1">
                        <label className="text-[10px] uppercase font-bold text-slate-500">Phép so sánh:</label>
                        <select
                          value={newThenRule.op}
                          onChange={(e) => setNewThenRule({ ...newThenRule, op: e.target.value })}
                          className="w-full bg-slate-50 border border-slate-300 rounded px-2 py-1 text-xs text-slate-800 font-medium"
                        >
                          <option value="==">Bằng (==)</option>
                          <option value="!=">Khác (!=)</option>
                          <option value=">">Lớn hơn (&gt;)</option>
                          <option value="<">Nhỏ hơn (&lt;)</option>
                          <option value=">=">Lớn hơn hoặc bằng (&gt;=)</option>
                          <option value="<=">Nhỏ hơn hoặc bằng (&lt;=)</option>
                          <option value="chứa">Chứa từ (chứa)</option>
                          <option value="không chứa">Không chứa từ</option>
                          <option value="trống">Để trống (rỗng)</option>
                          <option value="không trống">Không để trống</option>
                        </select>
                      </div>
                    </div>

                    {newThenRule.op !== "trống" && newThenRule.op !== "không trống" && (
                      <div className="space-y-2">
                        <div className="flex items-center gap-1.5">
                          <input
                            type="checkbox"
                            id="newThenRule_isFieldCompare"
                            checked={newThenRule.isFieldCompare}
                            onChange={(e) => setNewThenRule({ ...newThenRule, isFieldCompare: e.target.checked, val: "" })}
                            className="cursor-pointer"
                          />
                          <label htmlFor="newThenRule_isFieldCompare" className="text-[11px] text-slate-600 font-bold cursor-pointer">So sánh với giá trị cột khác</label>
                        </div>

                        <div className="space-y-1">
                          <label className="text-[10px] uppercase font-bold text-slate-500">Giá trị so sánh:</label>
                          {newThenRule.isFieldCompare ? (
                            <select
                              value={newThenRule.val}
                              onChange={(e) => setNewThenRule({ ...newThenRule, val: e.target.value })}
                              className="w-full bg-slate-50 border border-slate-300 rounded px-2 py-1 text-xs text-slate-800 font-medium"
                            >
                              <option value="">-- Chọn cột để so sánh --</option>
                              {columns.map(col => (
                                <option key={col} value={col}>{col}</option>
                              ))}
                            </select>
                          ) : (
                            <input
                              type="text"
                              value={newThenRule.val}
                              onChange={(e) => setNewThenRule({ ...newThenRule, val: e.target.value })}
                              placeholder="Nhập giá trị chữ hoặc số..."
                              className="w-full bg-slate-50 border border-slate-300 rounded px-2 py-1 text-xs text-slate-800"
                            />
                          )}
                        </div>
                      </div>
                    )}

                    <button
                      onClick={() => handleLogicRuleAdd("then")}
                      className="w-full bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-xs py-1.5 rounded border border-indigo-200 cursor-pointer transition-all flex items-center justify-center gap-1"
                    >
                      ➕ Thêm điều kiện THÌ PHẢI
                    </button>
                  </div>

                  {/* Danh sách quy tắc THÌ PHẢI đã thêm */}
                  <div className="space-y-1.5">
                    <span className="text-[10.5px] font-bold text-slate-500 block uppercase">Ràng buộc đang áp dụng:</span>
                    {thenRules.length === 0 ? (
                      <div className="text-[11px] text-slate-400 italic bg-white border border-slate-150 p-3 rounded text-center">
                        Chưa thiết lập ràng buộc. Chỉ rà quét theo bộ lọc NẾU độc lập.
                      </div>
                    ) : (
                      <div className="space-y-1 max-h-[140px] overflow-y-auto">
                        {thenRules.map((rule, idx) => (
                          <div key={idx} className="bg-white border border-slate-200 px-2.5 py-1.5 rounded flex items-center justify-between text-xs font-mono">
                            <span className="text-slate-700">
                              [{rule.col}] <strong className="text-emerald-600">{rule.op}</strong> {rule.isFieldCompare ? `Cột [${rule.val}]` : `'${rule.val || "rỗng"}'`}
                            </span>
                            <button
                              onClick={() => setThenRules(thenRules.filter((_, i) => i !== idx))}
                              className="text-slate-400 hover:text-red-500 p-0.5 rounded hover:bg-slate-100 transition-colors cursor-pointer"
                            >
                              ✕
                            </button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* KHU VỰC THIẾT LẬP PHƯƠNG THỨC HOẠT ĐỘNG VÀ BẤM CHẠY */}
              <div className="bg-slate-50 rounded-xl p-5 border border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-5">
                <div className="space-y-2 max-w-xl">
                  <span className="text-[10.5px] font-bold text-slate-500 block uppercase font-mono">Phương thức kiểm chứng logic:</span>
                  <div className="flex flex-col sm:flex-row gap-4">
                    <label className="flex items-start gap-2 cursor-pointer">
                      <input
                        type="radio"
                        name="logicRuleMode"
                        value="must_satisfy"
                        checked={logicRuleMode === "must_satisfy"}
                        onChange={() => setLogicRuleMode("must_satisfy")}
                        className="mt-1"
                      />
                      <div>
                        <span className="text-xs font-bold text-slate-800 block">Ràng buộc logic (Must-Satisfy)</span>
                        <span className="text-[10.5px] text-slate-500 block leading-tight">Báo lỗi nếu thỏa mãn vế NẾU nhưng KHÔNG thỏa mãn vế THÌ PHẢI.</span>
                      </div>
                    </label>

                    <label className="flex items-start gap-2 cursor-pointer">
                      <input
                        type="radio"
                        name="logicRuleMode"
                        value="conflict"
                        checked={logicRuleMode === "conflict"}
                        onChange={() => setLogicRuleMode("conflict")}
                        className="mt-1"
                      />
                      <div>
                        <span className="text-xs font-bold text-slate-800 block">Mâu thuẫn logic (Conflict)</span>
                        <span className="text-[10.5px] text-slate-500 block leading-tight">Báo lỗi nếu thỏa mãn đồng thời CẢ HAI vế NẾU và THÌ PHẢI.</span>
                      </div>
                    </label>
                  </div>
                </div>

                <div className="flex items-center gap-3 shrink-0">
                  <button
                    onClick={() => {
                      setIfRules([]);
                      setThenRules([]);
                      setAiScanMetrics(null);
                    }}
                    className="bg-white hover:bg-slate-100 text-slate-700 font-bold text-xs py-2.5 px-4 rounded-xl border border-slate-300 cursor-pointer transition-all active:scale-95 flex items-center gap-1"
                  >
                    🗑️ Xoá Sạch Quy Tắc
                  </button>

                  <button
                    onClick={handleLogicCheck}
                    className="bg-gradient-to-r from-emerald-500 to-emerald-600 hover:from-emerald-600 hover:to-emerald-700 text-white font-bold text-xs py-2.5 px-6 rounded-xl border-0 cursor-pointer transition-all shadow-md active:scale-95 flex items-center gap-1.5"
                  >
                    ⚡ CHẠY KIỂM TRA LOGIC THỦ CÔNG
                  </button>
                </div>
              </div>
            </div>

            {/* PHÂN HỆ AI: HỌC LỆNH VÀ RÀ SOÁT THÔNG MINH QUA AI GEMINI */}
            <div className="bg-slate-50 border border-slate-200 rounded-2xl p-5 space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-200">
                <div>
                  <h4 className="text-sm font-extrabold text-indigo-600 flex items-center gap-2">
                    <Sparkles className="w-4 h-4 text-indigo-500 animate-pulse" /> 🚀 TỰ HỌC LỆNH PHÂN TÍCH VÀ RÀ QUÉT BẰNG AI
                  </h4>
                  <p className="text-[11px] text-slate-500 mt-0.5">
                    Gõ điều kiện lỗi bằng tiếng Việt tự nhiên. Trợ lý AI sẽ tự động biên dịch sang biểu thức máy tính để quét toàn bộ file gốc.
                  </p>
                </div>
                
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={handleExportAiRules}
                    className="bg-white hover:bg-indigo-50 text-indigo-700 border border-slate-200 px-2.5 py-1 rounded-lg text-[11px] font-semibold cursor-pointer transition-all flex items-center gap-1 active:scale-95"
                    title="Tải bộ quy tắc logic đã lưu về máy dưới dạng JSON"
                  >
                    <Download className="w-3" /> Xuất bộ nhớ luật (.json)
                  </button>
                  
                  <label
                    className="bg-white hover:bg-purple-50 text-indigo-700 border border-slate-200 px-2.5 py-1 rounded-lg text-[11px] font-semibold cursor-pointer transition-all flex items-center gap-1 active:scale-95"
                    title="Nạp bộ quy tắc logic (.json) từ máy tính của bạn"
                  >
                    <Upload className="w-3" /> Nạp tệp cấu hình
                    <input
                      type="file"
                      accept=".json"
                      onChange={handleImportAiRules}
                      className="hidden"
                    />
                  </label>
                </div>
              </div>

              {/* KHU VỰC NHẬP LỆNH AI */}
              <div className="grid grid-cols-1 md:grid-cols-4 gap-3 items-end">
                <div className="md:col-span-3 space-y-1.5">
                  <div className="flex items-center justify-between mb-1 flex-wrap gap-2">
                    <label className="text-xs font-bold text-slate-700 flex items-center gap-1">
                      ✍️ Nhập điều kiện lỗi tiếng Việt (Ví dụ: 'Tìm dòng có DonGia &lt; 0'):
                    </label>
                    <button
                      type="button"
                      onClick={toggleMicRule}
                      className={`flex items-center gap-1 text-[10px] font-extrabold px-2 py-0.5 rounded border transition-all cursor-pointer ${
                        isRecordingRuleMic 
                          ? "bg-rose-100 hover:bg-rose-200 text-rose-700 border-rose-300 animate-pulse font-bold" 
                          : "bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border-emerald-200"
                      }`}
                      title="Bấm để nói bằng tiếng Việt"
                    >
                      {isRecordingRuleMic ? (
                        <>
                          <MicOff className="w-3 h-3 text-rose-600 shrink-0" /> Dừng nghe
                        </>
                      ) : (
                        <>
                          <Mic className="w-3 h-3 text-emerald-600 shrink-0" /> Ghi âm (Nói)
                        </>
                      )}
                    </button>
                  </div>
                  <div className="relative">
                    <input
                      type="text"
                      value={aiRulePrompt}
                      onChange={(e) => setAiRulePrompt(e.target.value)}
                      placeholder="Nhập khẩu lệnh bằng tiếng Việt tự nhiên tại đây..."
                      className="w-full bg-white border border-slate-300 rounded-xl pl-3 pr-10 py-2 text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500 placeholder-slate-400"
                    />
                    {aiRulePrompt && (
                      <button
                        onClick={() => setAiRulePrompt("")}
                        className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs cursor-pointer"
                      >
                        ✕
                      </button>
                    )}
                  </div>
                </div>

                <div>
                  <button
                    onClick={() => handleAiLogicScan()}
                    className="w-full bg-indigo-600 hover:bg-indigo-700 active:scale-[0.98] text-white font-bold text-xs py-2 rounded-xl transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-md border-0"
                  >
                    <Sparkles className="w-3.5 h-3.5 text-amber-200" /> QUÉT BẰNG AI GEMINI
                  </button>
                </div>
              </div>

              {/* KẾT QUẢ DỊCH CỦA AI & KHU VỰC LƯU TRỮ */}
              {aiTranslatedExpression && (
                <div className="bg-white rounded-xl p-4 border border-slate-200 space-y-3 animate-fade-in text-xs shadow-sm">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                    <div className="font-mono text-[11px] text-slate-700">
                      <span className="text-indigo-600 font-bold">🤖 Biểu thức máy hiểu (JS):</span> <code className="bg-slate-50 px-2 py-0.5 rounded text-indigo-700 border border-slate-250 break-all">{aiTranslatedExpression}</code>
                    </div>
                    <span className="bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-0.5 rounded text-[10px] font-mono shrink-0 font-bold">Dịch thành công!</span>
                  </div>

                  <div className="bg-slate-50 p-3 rounded-lg border border-slate-200 flex flex-col sm:flex-row sm:items-center gap-3">
                    <div className="flex-1 space-y-1">
                      <span className="text-[10px] uppercase font-bold text-slate-500">🏷️ Đặt tên quy tắc này để lưu nhanh:</span>
                      <input
                        type="text"
                        value={customRuleName}
                        onChange={(e) => setCustomRuleName(e.target.value)}
                        placeholder="Ví dụ: Kiểm tra Đơn Giá âm, Mã ngành rỗng..."
                        className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-800"
                      />
                    </div>
                    <button
                      onClick={handleSaveAiRule}
                      className="bg-emerald-500 hover:bg-emerald-600 active:scale-[0.98] text-white font-bold text-xs px-4 py-2 rounded-lg transition-all cursor-pointer flex items-center gap-1 shrink-0 justify-center border-0"
                    >
                      💾 Lưu học lệnh thông minh
                    </button>
                  </div>
                </div>
              )}

              {/* BẢNG TỔNG HỢP KẾT QUẢ RÀ QUÉT AI */}
              {aiScanMetrics && (
                <div id="ai-scan-summary-section" className="bg-white rounded-2xl p-5 border border-slate-200 space-y-4 animate-fade-in text-xs shadow-md">
                  <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                    <h5 className="text-[12px] font-extrabold text-indigo-600 flex items-center gap-1.5 font-mono">
                      📊 BẢNG TỔNG HỢP KẾT QUẢ RÀ QUÉT AI
                    </h5>
                    <span className="text-[10px] bg-indigo-50 text-indigo-700 border border-indigo-200 px-2 py-0.5 rounded font-mono">
                      Khớp tự học lệnh của AI
                    </span>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div className="bg-slate-50 border border-slate-200 p-3.5 rounded-xl text-center space-y-1">
                      <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block">🔍 Tổng số đã quét</span>
                      <strong className="text-xl text-slate-900 font-mono font-bold">{aiScanMetrics.total.toLocaleString()}</strong>
                      <span className="text-[10px] text-slate-500 block">bản ghi dữ liệu gốc</span>
                    </div>

                    <div className={`border p-3.5 rounded-xl text-center space-y-1 transition-all ${
                      aiScanMetrics.violated > 0 
                        ? "bg-red-50 border-red-200 text-red-700" 
                        : "bg-slate-50 border-slate-200 text-slate-500"
                    }`}>
                      <span className="text-[10px] font-bold uppercase tracking-wider block">❌ Số dòng bị lỗi</span>
                      <strong className={`text-xl font-mono ${aiScanMetrics.violated > 0 ? "text-amber-600 font-bold" : "text-slate-500 font-normal"}`}>
                        {aiScanMetrics.violated.toLocaleString()}
                      </strong>
                      <span className="text-[10px] block opacity-80 text-amber-600">Tỷ lệ: {aiScanMetrics.violatedPercent}</span>
                    </div>

                    <div className="bg-emerald-50 border border-emerald-200 p-3.5 rounded-xl text-center space-y-1 text-emerald-700">
                      <span className="text-[10px] text-emerald-600 font-bold uppercase tracking-wider block">✅ Số dòng đạt chuẩn</span>
                      <strong className="text-xl text-emerald-600 font-mono font-bold">{aiScanMetrics.passed.toLocaleString()}</strong>
                      <span className="text-[10px] text-emerald-500 block">Tỷ lệ: {aiScanMetrics.passedPercent}</span>
                    </div>
                  </div>

                  {/* Bộ lọc hiển thị nhanh cho kết quả quét AI */}
                  <div className="bg-slate-50 px-4 py-2.5 rounded-xl border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-slate-800">
                    <div>
                      <p className="font-bold text-slate-900">ℹ️ Đang hiển thị trực quan dưới bảng:</p>
                      <p className="text-[11px] text-slate-500 mt-0.5 font-sans">
                        {logicFilterMode === "violated" 
                          ? `Chỉ hiển thị ${aiScanMetrics.violated} dòng bị phát hiện vi phạm (quy quét được tô nền lỗi đỏ).`
                          : `Đang hiển thị toàn bộ các dòng được rà quét (Bao gồm cả Đạt ✅ và Vi Phạm ❌).`}
                      </p>
                    </div>
                    <div className="flex gap-2 shrink-0">
                      <button
                        onClick={() => setLogicFilterMode("violated")}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all active:scale-95 cursor-pointer border ${
                          logicFilterMode === "violated" 
                            ? "bg-red-600 text-white border-red-550 shadow-sm" 
                            : "bg-white text-slate-600 hover:text-slate-900 border-slate-250"
                        }`}
                      >
                        ❌ Xem dòng Lỗi ({aiScanMetrics.violated})
                      </button>
                      <button
                        onClick={() => setLogicFilterMode("all")}
                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all active:scale-95 cursor-pointer border ${
                          logicFilterMode === "all" 
                            ? "bg-indigo-600 text-white border-indigo-550 shadow-sm" 
                            : "bg-white text-slate-600 hover:text-slate-900 border-slate-250"
                        }`}
                      >
                        🌐 Xem tất cả ({aiScanMetrics.total})
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* DANH SÁCH CÁC QUY TẮC ĐÃ LƯU (HỌC LỆNH CHẠY NHANH) */}
              {savedAiRules.length > 0 && (
                <div className="space-y-2 pt-2 border-t border-slate-200">
                  <div className="text-[10px] font-bold text-indigo-600 tracking-wider font-mono uppercase">⚡ BỘ NHỚ HỌC LỆNH THÔNG MINH (BẤM NÚT CHẠY LUÔN KHÔNG CẦN CHỜ DỊCH AI):</div>
                  <div className="flex flex-wrap gap-2">
                    {savedAiRules.map(rule => (
                      <button
                        key={rule.id}
                        onClick={() => handleAiLogicScan(rule.prompt)}
                        className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 px-2.5 py-1.5 rounded-lg text-[10px] font-semibold cursor-pointer transition-all flex items-center gap-1"
                        title={rule.prompt}
                      >
                        ⚡ {rule.name}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {/* BẢNG TRỰC QUAN HIỂN THỊ KẾT QUẢ QUÉT LOGIC */}
              {filteredLogicData.length > 0 && (
                <div className="pt-6 border-t border-slate-200">
                  <MainDataInlinePreview 
                    data={filteredLogicData} 
                    columns={columns.includes("Loi_Logic") ? columns : ["Loi_Logic", ...columns]} 
                    title="BẢNG DÒNG DỮ LIỆU ĐÃ KIỂM TRA LOGIC" 
                    subtitle={`Đang hiển thị ${filteredLogicData.length} dòng thuộc bộ lọc "${logicFilterMode === "violated" ? "Dòng lỗi (vi phạm)" : logicFilterMode === "if_satisfied" ? "Thỏa mãn vế NẾU" : "Tất cả các dòng liên quan đã quét"}"`}
                    mapping={mapping}
                    onExportExcel={onExportExcel}
                  />
                </div>
              )}

            </div>
          </div>
        ) : (
          <div className="p-8 text-center text-slate-400 bg-slate-50 border border-dashed border-slate-300 rounded-xl">
            Chưa có dữ liệu nguồn. Vui lòng nạp file ở trang chủ trước.
          </div>
        )}
      </div>
    </div>
  );
}
