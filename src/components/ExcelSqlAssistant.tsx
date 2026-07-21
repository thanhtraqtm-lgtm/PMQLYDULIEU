import React, { useState } from "react";
import { 
  FileSpreadsheet, 
  Database, 
  Terminal, 
  Copy, 
  Check, 
  Sparkles, 
  HelpCircle, 
  ArrowRight, 
  Send, 
  RefreshCw, 
  Code,
  BookOpen,
  Settings,
  ChevronRight
} from "lucide-react";
import { GoogleGenAI } from "@google/genai";

interface ExcelSqlAssistantProps {
  mainData?: any[];
  fileName?: string;
}

export default function ExcelSqlAssistant({ mainData = [], fileName = "Dữ liệu chính" }: ExcelSqlAssistantProps) {
  const [activeSubTab, setActiveSubTab] = useState<"excel" | "vba" | "sql" | "chatbot">("excel");
  const [userPrompt, setUserPrompt] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);
  const [generatedCode, setGeneratedCode] = useState("");
  const [generatedExplanation, setGeneratedExplanation] = useState("");
  const [copied, setCopied] = useState(false);
  const [vietnameseExcelFormat, setVietnameseExcelFormat] = useState(false); // Dùng dấu ";" thay cho "," trong công thức

  // Chatbot state
  const [chatMessages, setChatMessages] = useState<Array<{ sender: "user" | "ai"; text: string; code?: string }>>([
    {
      sender: "ai",
      text: "Chào bạn! Tôi là trợ lý AI chuyên về công thức Excel, VBA và truy vấn SQL dữ liệu. Bạn cần tôi viết công thức gì hay truy vấn mẫu biểu nào hôm nay?"
    }
  ]);
  const [chatInput, setChatInput] = useState("");
  const [chatLoading, setChatLoading] = useState(false);

  // Mẫu gợi ý (Presets)
  const excelPresets = [
    {
      title: "Tìm kiếm bằng XLOOKUP nâng cao",
      prompt: "Tìm tên ngành kinh doanh từ bảng Danh mục VSIC dựa vào cột Mã ngành kinh doanh, nếu không tìm thấy trả về 'Không tìm thấy'"
    },
    {
      title: "Cộng nhiều điều kiện SUMIFS",
      prompt: "Tính tổng doanh thu (cột DoanhThu) cho đơn vị có Mã địa bàn là 'X12' và mã Ngành cấp 1 là 'A'"
    },
    {
      title: "Trích xuất chuỗi bằng hàm Text",
      prompt: "Trích xuất 2 ký tự đầu tiên của mã cơ sở (cột MaCS) và ghép với mã tỉnh nếu cột Tỉnh không rỗng"
    },
    {
      title: "Kiểm tra mã số thuế hợp lệ",
      prompt: "Kiểm tra độ dài cột MaSoThue nếu bằng 10 hoặc 13 chữ số thì đúng, ngược lại báo sai bằng hàm IF và LEN"
    }
  ];

  const vbaPresets = [
    {
      title: "Tự động định dạng bảng báo cáo",
      prompt: "Viết macro VBA tự động kẻ bảng, bôi đậm dòng tiêu đề thứ nhất, đặt màu nền tiêu đề là xanh lam nhạt, căn lề giữa và tự động dãn cột cho bảng tính đang chọn"
    },
    {
      title: "Tách trang tính theo mã đơn vị",
      prompt: "Viết VBA duyệt qua cột 'Mã Đơn Vị' ở Sheet1, tạo các Sheet mới theo từng mã đơn vị duy nhất và sao chép dữ liệu tương ứng sang Sheet đó"
    },
    {
      title: "Lọc bỏ dữ liệu lỗi (Error values)",
      prompt: "Viết VBA quét qua vùng dữ liệu đang chọn, thay thế toàn bộ giá trị lỗi như #N/A, #DIV/0!, #VALUE! thành ô rỗng"
    }
  ];

  const sqlPresets = [
    {
      title: "Kết nối dữ liệu với danh mục VSIC",
      prompt: "Viết truy vấn SQL kết hợp bảng 'DoanhNghiep' và bảng 'DanhMucVSIC' dựa trên cột 'MaNganh'. Lấy ra Tên doanh nghiệp, Mã ngành và Tên ngành"
    },
    {
      title: "Thống kê số lượng DN theo địa bàn",
      prompt: "Viết truy vấn SQL đếm số lượng doanh nghiệp và tính tổng doanh thu theo từng 'MaXa', sắp xếp theo số lượng doanh nghiệp giảm dần"
    },
    {
      title: "Tìm kiếm bản ghi trùng lặp MST",
      prompt: "Viết truy vấn SQL tìm các bản ghi bị trùng lặp cột 'MaSoThue' (xuất hiện lớn hơn 1 lần) trong bảng 'DanhSachKhaoSat'"
    }
  ];

  const handleCopy = (text: string) => {
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const formatExcelFormula = (formula: string, useSemicolon: boolean) => {
    if (!formula) return "";
    if (useSemicolon) {
      // Thay đổi dấu phẩy ngăn cách tham số thành dấu chấm phẩy (phổ biến ở cài đặt Windows Việt Nam)
      // Ví dụ: =SUMIFS(A:A, B:B, "X") -> =SUMIFS(A:A; B:B; "X")
      // Đây là một tiện ích cực kì hữu hiệu thực tế!
      let inQuote = false;
      let result = "";
      for (let i = 0; i < formula.length; i++) {
        const char = formula[i];
        if (char === '"') inQuote = !inQuote;
        if (char === ',' && !inQuote) {
          result += ";";
        } else {
          result += char;
        }
      }
      return result;
    }
    return formula;
  };

  const getApiKey = () => {
    return (import.meta as any).env?.VITE_GEMINI_API_KEY || "";
  };

  const generateResponse = async (type: "excel" | "vba" | "sql", promptText: string) => {
    const key = getApiKey();
    if (!key) {
      // Cung cấp giải pháp fallback thông minh nếu không có API key
      const mockResponses: Record<string, { code: string; exp: string }> = {
        excel: {
          code: '=XLOOKUP(A2; DanhMucVSIC!$A$2:$A$1000; DanhMucVSIC!$B$2:$B$1000; "Không tìm thấy"; 0)',
          exp: "Giải thích công thức:\n1. A2: Giá trị mã ngành cần tìm kiếm.\n2. DanhMucVSIC!$A$2:$A$1000: Vùng chứa mã ngành trong bảng danh mục gốc.\n3. DanhMucVSIC!$B$2:$B$1000: Vùng chứa tên ngành tương ứng cần trả về.\n4. \"Không tìm thấy\": Giá trị hiển thị dự phòng nếu không tìm thấy mã ngành.\n5. 0: Chế độ khớp chính xác tuyệt đối."
        },
        vba: {
          code: `Sub FormatReportTable()\n    Dim ws As Worksheet\n    Set ws = ActiveSheet\n    Dim lastRow As Long, lastCol As Long\n    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\n    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column\n    \n    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))\n        .Borders.LineStyle = xlContinuous\n        .Borders.Weight = xlThin\n    End With\n    \n    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))\n        .Font.Bold = True\n        .Interior.Color = RGB(220, 230, 242)\n        .HorizontalAlignment = xlCenter\n    End With\n    ws.Columns.AutoFit\n    MsgBox "Đã định dạng bảng thành công!", vbInformation\nEnd Sub`,
          exp: "Giải thích code VBA:\n- Xác định tự động dòng cuối (lastRow) và cột cuối (lastCol).\n- Kẻ đường viền mảnh (.Borders.Weight = xlThin) cho toàn bộ bảng dữ liệu.\n- Thiết lập dòng tiêu đề (Dòng 1): bôi đậm, tô nền xanh dương nhạt (RGB 220, 230, 242), căn giữa.\n- AutoFit tự động co dãn cột theo nội dung dài nhất."
        },
        sql: {
          code: `SELECT \n    dn.TenDoanhNghiep,\n    dn.MaNganh,\n    vsic.TenNganh\nFROM DoanhNghiep dn\nINNER JOIN DanhMucVSIC vsic ON dn.MaNganh = vsic.MaNganh;`,
          exp: "Giải thích truy vấn SQL:\n- SELECT: Lọc ra 3 trường dữ liệu cần thiết.\n- FROM DoanhNghiep dn: Chỉ định bảng chính chứa thông tin doanh nghiệp (đặt bí danh là dn).\n- INNER JOIN ... ON: Kết nối vật lý với bảng danh mục ngành nghề VSIC (bí danh vsic) dựa trên mối quan hệ trùng khớp mã ngành giữa 2 bảng."
        }
      };

      setIsGenerating(true);
      setTimeout(() => {
        const fallback = mockResponses[type];
        setGeneratedCode(fallback.code);
        setGeneratedExplanation(fallback.exp + "\n\n💡 (Lưu ý: Đây là phản hồi mẫu vì chưa cấu hình khóa API VITE_GEMINI_API_KEY)");
        setIsGenerating(false);
      }, 800);
      return;
    }

    setIsGenerating(false);
    setIsGenerating(true);
    setGeneratedCode("");
    setGeneratedExplanation("");

    try {
      const ai = new GoogleGenAI({
        apiKey: key,
        httpOptions: { headers: { "User-Agent": "aistudio-build" } }
      });

      const sysInstruction = `Bạn là Trợ lý AI cao cấp chuyên viết công thức Excel, mã VBA và truy vấn SQL tối ưu. 
Hãy trả về câu trả lời bằng tiếng Việt, được định dạng theo cấu trúc:
1. Đặt đoạn CODE hoặc CÔNG THỨC trong một khối mã Markdown duy nhất (để người dùng dễ sao chép). Không sử dụng nhiều khối mã, chỉ dùng đúng 1 khối mã chứa công thức hoặc script chính.
2. Trực tiếp đưa ra lời giải thích chi tiết, ngắn gọn, dễ hiểu bên dưới khối mã.
3. Luôn bám sát thực tế hành chính và thống kê của Việt Nam.`;

      let prompt = "";
      if (type === "excel") {
        prompt = `Hãy viết công thức Excel cho yêu cầu sau: "${promptText}". 
Lưu ý: Trả về công thức Excel chuẩn bắt đầu bằng dấu "=" bằng tiếng Anh (ví dụ: VLOOKUP, SUMIFS, IF, XLOOKUP, INDEX, MATCH).
Nếu cần, hãy hướng dẫn cả cách áp dụng trên dải ô.`;
      } else if (type === "vba") {
        prompt = `Hãy viết mã VBA Macro Excel cho yêu cầu sau: "${promptText}".
Đảm bảo mã có đầy đủ khai báo biến (Dim), xử lý lỗi cơ bản nếu cần và ghi chú giải thích cụ thể bằng tiếng Việt.`;
      } else {
        prompt = `Hãy viết câu lệnh truy vấn SQL cho yêu cầu sau: "${promptText}".
Đảm bảo cú pháp SQL chuẩn (ANSI SQL), có thụt lề rõ ràng, trực quan.`;
      }

      const response = await ai.models.generateContent({
        model: "gemini-3.6-flash",
        contents: prompt,
        config: {
          systemInstruction: sysInstruction,
          temperature: 0.2
        }
      });

      const text = response.text || "";
      
      // Phân tách khối mã và giải thích
      const codeBlockRegex = /```[a-zA-Z]*\n([\s\S]*?)```/g;
      const match = codeBlockRegex.exec(text);
      if (match) {
        setGeneratedCode(match[1].trim());
        setGeneratedExplanation(text.replace(match[0], "").trim());
      } else {
        // Fallback nếu không có block markdown code
        setGeneratedCode(text.split("\n")[0] || "");
        setGeneratedExplanation(text);
      }
    } catch (err: any) {
      console.error("Lỗi Gemini:", err);
      setGeneratedExplanation(`Lỗi khi tạo phản hồi từ AI: ${err.message}. Vui lòng thử lại.`);
    } finally {
      setIsGenerating(false);
    }
  };

  const handleSendChatMessage = async () => {
    if (!chatInput.trim()) return;
    const userMsg = chatInput.trim();
    setChatMessages(prev => [...prev, { sender: "user", text: userMsg }]);
    setChatInput("");
    setChatLoading(true);

    const key = getApiKey();
    if (!key) {
      setTimeout(() => {
        setChatMessages(prev => [
          ...prev,
          { 
            sender: "ai", 
            text: `Bạn vừa hỏi: "${userMsg}". Hiện tại chưa có khóa API VITE_GEMINI_API_KEY, vui lòng bổ sung để tôi có thể suy luận trực tiếp dựa trên ngữ cảnh thực tế dữ liệu của bạn nhé!` 
          }
        ]);
        setChatLoading(false);
      }, 600);
      return;
    }

    try {
      const ai = new GoogleGenAI({
        apiKey: key,
        httpOptions: { headers: { "User-Agent": "aistudio-build" } }
      });

      // Tạo chuỗi lịch sử chat
      const historyContext = chatMessages.slice(-5).map(m => `${m.sender === "user" ? "User" : "AI"}: ${m.text}`).join("\n");

      const sysInstruction = `Bạn là Trợ lý AI hỗ trợ phân tích số liệu thống kê liên ngành, chuyên nghiệp về Excel, SQL, VBA, Python và R.
Hãy trả lời các thắc mắc về kỹ thuật dữ liệu, lọc trùng, chuẩn hóa ngành nghề VSIC của Việt Nam một cách chính xác, thân thiện và hữu ích.`;

      const response = await ai.models.generateContent({
        model: "gemini-3.6-flash",
        contents: `Lịch sử hội thoại gần đây:\n${historyContext}\n\nCâu hỏi mới nhất của người dùng: "${userMsg}"\nHãy trả lời bằng tiếng Việt chuẩn xác nhất.`,
        config: {
          systemInstruction: sysInstruction,
          temperature: 0.5
        }
      });

      const aiText = response.text || "Tôi không nhận được câu trả lời từ máy chủ AI.";
      setChatMessages(prev => [...prev, { sender: "ai", text: aiText }]);
    } catch (err: any) {
      setChatMessages(prev => [...prev, { sender: "ai", text: `Lỗi kết nối máy chủ AI: ${err.message}` }]);
    } finally {
      setChatLoading(false);
    }
  };

  return (
    <div className="bg-white text-slate-800 rounded-2xl border border-slate-200/80 shadow-md overflow-hidden animate-fade-in" id="excel_sql_assistant_root">
      
      {/* HEADER CỦA TIỆN ÍCH */}
      <div className="bg-gradient-to-r from-slate-50 via-indigo-50/40 to-slate-50 px-6 py-5 border-b border-slate-200/80 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600">
            <Sparkles className="w-5 h-5 animate-pulse" />
          </div>
          <div>
            <h2 className="text-base font-extrabold tracking-wide uppercase font-sans text-slate-900">
              TRỢ LÝ CÔNG THỨC EXCEL &amp; TRUY VẤN SQL THÔNG MINH
            </h2>
            <p className="text-xs text-slate-500 mt-0.5 font-medium">
              Tạo công thức, viết VBA Macro và sinh truy vấn cơ sở dữ liệu tốc độ cao bằng trí tuệ nhân tạo.
            </p>
          </div>
        </div>
        
        {/* THÔNG TIN FILE NGUỒN HIỆN TẠI ĐỂ SỬ DỤNG CHO AI */}
        {mainData && mainData.length > 0 && (
          <div className="bg-indigo-50 border border-indigo-100 rounded-xl px-3 py-1.5 flex items-center gap-2">
            <div className="w-2 h-2 rounded-full bg-emerald-500 animate-ping" />
            <span className="text-[10px] font-mono text-indigo-700 font-bold">
              KẾT NỐI: {fileName} ({mainData.length} dòng)
            </span>
          </div>
        )}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 min-h-[580px]">
        
        {/* CỘT TRÁI: ĐIỀU HƯỚNG TABS & PRESETS */}
        <div className="lg:col-span-4 border-r border-slate-200 bg-slate-50/50 p-5 flex flex-col justify-between">
          <div className="space-y-6">
            
            {/* CHỌN CÔNG CỤ */}
            <div className="space-y-2">
              <span className="text-[10px] font-bold text-indigo-600 block uppercase tracking-wider font-sans">
                🛠️ Lựa chọn công cụ hỗ trợ
              </span>
              <div className="grid grid-cols-2 gap-2">
                <button
                  onClick={() => setActiveSubTab("excel")}
                  className={`p-3 rounded-xl border text-xs font-bold flex flex-col items-center justify-center gap-2 transition cursor-pointer ${
                    activeSubTab === "excel"
                      ? "bg-indigo-50 border-indigo-200 text-indigo-700 shadow-sm"
                      : "bg-white border-slate-200 text-slate-600 hover:border-slate-300 hover:text-slate-900"
                  }`}
                >
                  <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                  Công thức Excel
                </button>
                <button
                  onClick={() => setActiveSubTab("vba")}
                  className={`p-3 rounded-xl border text-xs font-bold flex flex-col items-center justify-center gap-2 transition cursor-pointer ${
                    activeSubTab === "vba"
                      ? "bg-indigo-50 border-indigo-200 text-indigo-700 shadow-sm"
                      : "bg-white border-slate-200 text-slate-600 hover:border-slate-300 hover:text-slate-900"
                  }`}
                >
                  <Terminal className="w-4 h-4 text-amber-600" />
                  VBA Macro
                </button>
                <button
                  onClick={() => setActiveSubTab("sql")}
                  className={`p-3 rounded-xl border text-xs font-bold flex flex-col items-center justify-center gap-2 transition cursor-pointer ${
                    activeSubTab === "sql"
                      ? "bg-indigo-50 border-indigo-200 text-indigo-700 shadow-sm"
                      : "bg-white border-slate-200 text-slate-600 hover:border-slate-300 hover:text-slate-900"
                  }`}
                >
                  <Database className="w-4 h-4 text-indigo-600" />
                  Truy vấn SQL
                </button>
                <button
                  onClick={() => setActiveSubTab("chatbot")}
                  className={`p-3 rounded-xl border text-xs font-bold flex flex-col items-center justify-center gap-2 transition cursor-pointer ${
                    activeSubTab === "chatbot"
                      ? "bg-indigo-50 border-indigo-200 text-indigo-700 shadow-sm"
                      : "bg-white border-slate-200 text-slate-600 hover:border-slate-300 hover:text-slate-900"
                  }`}
                >
                  <Sparkles className="w-4 h-4 text-purple-600" />
                  Trò chuyện AI
                </button>
              </div>
            </div>

            {/* DANH SÁCH GỢI Ý NHANH (Chỉ hiển thị khi không ở tab chatbot) */}
            {activeSubTab !== "chatbot" && (
              <div className="space-y-2 animate-fade-in">
                <span className="text-[10px] font-bold text-slate-500 block uppercase tracking-wider font-sans">
                  💡 Gợi ý tình huống thực tế
                </span>
                <div className="space-y-1.5 max-h-[300px] overflow-y-auto pr-1">
                  {activeSubTab === "excel" && excelPresets.map((item, idx) => (
                    <button
                      key={idx}
                      onClick={() => setUserPrompt(item.prompt)}
                      className="w-full text-left p-2.5 rounded-lg bg-white border border-slate-200 hover:border-indigo-500/40 hover:bg-slate-50 text-[11px] font-medium transition cursor-pointer flex gap-2 items-start shadow-sm"
                    >
                      <ChevronRight className="w-3.5 h-3.5 text-indigo-500 shrink-0 mt-0.5" />
                      <div>
                        <p className="font-bold text-slate-700">{item.title}</p>
                        <p className="text-[10px] text-slate-400 line-clamp-1 mt-0.5">{item.prompt}</p>
                      </div>
                    </button>
                  ))}

                  {activeSubTab === "vba" && vbaPresets.map((item, idx) => (
                    <button
                      key={idx}
                      onClick={() => setUserPrompt(item.prompt)}
                      className="w-full text-left p-2.5 rounded-lg bg-white border border-slate-200 hover:border-indigo-500/40 hover:bg-slate-50 text-[11px] font-medium transition cursor-pointer flex gap-2 items-start shadow-sm"
                    >
                      <ChevronRight className="w-3.5 h-3.5 text-amber-500 shrink-0 mt-0.5" />
                      <div>
                        <p className="font-bold text-slate-700">{item.title}</p>
                        <p className="text-[10px] text-slate-400 line-clamp-1 mt-0.5">{item.prompt}</p>
                      </div>
                    </button>
                  ))}

                  {activeSubTab === "sql" && sqlPresets.map((item, idx) => (
                    <button
                      key={idx}
                      onClick={() => setUserPrompt(item.prompt)}
                      className="w-full text-left p-2.5 rounded-lg bg-white border border-slate-200 hover:border-indigo-500/40 hover:bg-slate-50 text-[11px] font-medium transition cursor-pointer flex gap-2 items-start shadow-sm"
                    >
                      <ChevronRight className="w-3.5 h-3.5 text-emerald-500 shrink-0 mt-0.5" />
                      <div>
                        <p className="font-bold text-slate-700">{item.title}</p>
                        <p className="text-[10px] text-slate-400 line-clamp-1 mt-0.5">{item.prompt}</p>
                      </div>
                    </button>
                  ))}
                </div>
              </div>
            )}

          </div>

          <div className="mt-6 pt-4 border-t border-slate-200 text-[10px] text-slate-500 leading-relaxed font-sans">
            📌 <strong>Mẹo nhỏ:</strong> Bạn có thể mô tả bằng ngôn ngữ tự nhiên hàng ngày. Ví dụ: <i>"Nếu cột A bằng rỗng và cột B lớn hơn 100 thì lấy cột C nhân 10%"</i>, AI sẽ viết chính xác cấu trúc hàm lồng nhau cho bạn!
          </div>
        </div>

        {/* CỘT PHẢI: KHU VỰC TẠO VÀ HIỂN THỊ KẾT QUẢ */}
        <div className="lg:col-span-8 p-6 bg-slate-50/30 flex flex-col">
          
          {activeSubTab !== "chatbot" ? (
            <div className="space-y-5 flex-1 flex flex-col justify-between animate-fade-in">
              
              {/* PHẦN NHẬP PHÁT BIỂU YÊU CẦU */}
              <div className="space-y-2.5">
                <div className="flex items-center justify-between">
                  <label className="text-xs font-bold text-indigo-700 block uppercase tracking-wider font-sans">
                    ✍️ Mô tả yêu cầu của bạn bằng tiếng Việt
                  </label>
                  
                  {activeSubTab === "excel" && (
                    <button 
                      onClick={() => setVietnameseExcelFormat(!vietnameseExcelFormat)}
                      className={`text-[10px] px-2.5 py-1 rounded-md border font-bold transition flex items-center gap-1 cursor-pointer ${
                        vietnameseExcelFormat 
                          ? "bg-emerald-50 border-emerald-200 text-emerald-700" 
                          : "bg-white border-slate-200 text-slate-500 hover:text-slate-800"
                      }`}
                      title="Chuyển đổi dấu ngăn cách đối số từ dấu phẩy sang dấu chấm phẩy"
                    >
                      Định dạng Windows VN (Dùng dấu ";") : {vietnameseExcelFormat ? "BẬT" : "TẮT"}
                    </button>
                  )}
                </div>

                <div className="relative">
                  <textarea
                    rows={4}
                    placeholder={
                      activeSubTab === "excel" 
                        ? "Ví dụ: Viết hàm tìm kiếm mã số thuế từ bảng phụ Sheet2 và điền vào cột MST ở Sheet1, nếu lỗi trả về ô trống..."
                        : activeSubTab === "vba"
                          ? "Ví dụ: Viết macro VBA tự động gộp dữ liệu từ 3 sheet có tên bắt đầu bằng 'Huyen' thành 1 bảng tổng hợp duy nhất..."
                          : "Ví dụ: Truy vấn đếm số doanh nghiệp có vốn điều lệ trên 10 tỷ đồng và gom nhóm theo ngành cấp 2..."
                    }
                    value={userPrompt}
                    onChange={(e) => setUserPrompt(e.target.value)}
                    className="w-full px-4 py-3 bg-white border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500/20 shadow-inner transition font-sans leading-relaxed resize-none"
                  />
                  <div className="absolute bottom-3 right-3 flex items-center gap-2">
                    {userPrompt && (
                      <button
                        onClick={() => setUserPrompt("")}
                        className="text-[10px] px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded font-bold transition cursor-pointer"
                      >
                        Xóa
                      </button>
                    )}
                    <button
                      disabled={isGenerating || !userPrompt.trim()}
                      onClick={() => generateResponse(activeSubTab, userPrompt)}
                      className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-500 disabled:bg-slate-100 disabled:text-slate-400 text-white text-xs font-bold rounded-lg transition flex items-center gap-1.5 shadow-md shadow-indigo-100 cursor-pointer"
                    >
                      {isGenerating ? (
                        <>
                          <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                          Đang tạo...
                        </>
                      ) : (
                        <>
                          <Sparkles className="w-3.5 h-3.5" />
                          Tạo ngay
                        </>
                      )}
                    </button>
                  </div>
                </div>
              </div>

              {/* BẢNG KẾT QUẢ ĐẦU RA */}
              <div className="flex-1 min-h-[250px] border border-slate-200 bg-white rounded-2xl p-5 flex flex-col justify-between shadow-sm">
                
                {isGenerating ? (
                  <div className="flex-1 flex flex-col items-center justify-center space-y-3 py-10">
                    <div className="w-10 h-10 rounded-full border-2 border-indigo-600 border-t-transparent animate-spin" />
                    <p className="text-xs text-slate-500 font-sans font-medium">
                      Trí tuệ nhân tạo đang phân tích cú pháp dữ liệu và lập trình mã nguồn...
                    </p>
                  </div>
                ) : generatedCode ? (
                  <div className="space-y-4 flex-1 flex flex-col justify-between animate-fade-in">
                    
                    <div className="space-y-2">
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] font-bold text-slate-500 block uppercase tracking-wider font-sans">
                          💾 MÃ NGUỒN / CÔNG THỨC KHUYÊN DÙNG
                        </span>
                        
                        <button
                          onClick={() => handleCopy(formatExcelFormula(generatedCode, activeSubTab === "excel" && vietnameseExcelFormat))}
                          className="px-3 py-1 bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-700 text-[10px] font-bold rounded-lg transition flex items-center gap-1.5 cursor-pointer"
                        >
                          {copied ? (
                            <>
                              <Check className="w-3 h-3 text-emerald-600" />
                              <span className="text-emerald-600">Đã sao chép!</span>
                            </>
                          ) : (
                            <>
                              <Copy className="w-3 h-3" />
                              Sao chép mã
                            </>
                          )}
                        </button>
                      </div>

                      <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl overflow-x-auto">
                        <pre className="text-xs font-mono text-slate-800 whitespace-pre-wrap select-all font-semibold">
                          {formatExcelFormula(generatedCode, activeSubTab === "excel" && vietnameseExcelFormat)}
                        </pre>
                      </div>
                    </div>

                    <div className="space-y-1.5 border-t border-slate-100 pt-3 flex-1 overflow-y-auto max-h-[160px] pr-1">
                      <span className="text-[10px] font-bold text-slate-500 block uppercase tracking-wider font-sans">
                        📖 HƯỚNG DẪN CHI TIẾT &amp; GIẢI THÍCH HÀM
                      </span>
                      <div className="text-[11px] text-slate-600 font-sans leading-relaxed whitespace-pre-line">
                        {generatedExplanation}
                      </div>
                    </div>

                  </div>
                ) : (
                  <div className="flex-1 flex flex-col items-center justify-center text-center space-y-4 py-10">
                    <div className="w-12 h-12 rounded-full bg-slate-50 border border-slate-100 flex items-center justify-center text-slate-400">
                      <Code className="w-6 h-6" />
                    </div>
                    <div>
                      <h4 className="text-xs font-bold text-slate-700 uppercase">Chưa có mã nguồn được tạo</h4>
                      <p className="text-[11px] text-slate-400 font-sans mt-1 max-w-sm font-medium">
                        Nhập yêu cầu của bạn ở ô phía trên hoặc nhấn vào một gợi ý nhanh ở danh sách bên trái để sinh mã nguồn tự động.
                      </p>
                    </div>
                  </div>
                )}

              </div>

            </div>
          ) : (
            
            /* TAB CHATBOT AI CHUYÊN SÂU */
            <div className="flex-1 flex flex-col justify-between h-full min-h-[500px] animate-fade-in">
              
              {/* KHU VỰC CUỘN TIN NHẮN */}
              <div className="flex-1 overflow-y-auto space-y-4 max-h-[380px] mb-4 pr-1 bg-slate-50/40 rounded-2xl border border-slate-200/60 p-4 shadow-inner">
                {chatMessages.map((msg, idx) => (
                  <div
                    key={idx}
                    className={`flex gap-3 max-w-[85%] ${msg.sender === "user" ? "ml-auto flex-row-reverse" : "mr-auto"}`}
                  >
                    <div className={`w-7 h-7 rounded-lg flex items-center justify-center text-xs font-bold shrink-0 ${
                      msg.sender === "user" ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-600 border border-slate-200"
                    }`}>
                      {msg.sender === "user" ? "U" : "AI"}
                    </div>
                    <div className={`p-3 rounded-2xl text-xs leading-relaxed font-sans ${
                      msg.sender === "user" 
                        ? "bg-indigo-600 text-white rounded-tr-none" 
                        : "bg-white border border-slate-200 text-slate-700 rounded-tl-none whitespace-pre-line shadow-sm"
                    }`}>
                      {msg.text}
                    </div>
                  </div>
                ))}
                {chatLoading && (
                  <div className="flex gap-3 mr-auto items-center">
                    <div className="w-7 h-7 rounded-lg bg-slate-100 text-slate-500 border border-slate-200 flex items-center justify-center text-xs font-bold animate-pulse">
                      AI
                    </div>
                    <div className="flex gap-1.2 p-3 rounded-2xl bg-white border border-slate-200 text-slate-400 rounded-tl-none text-xs">
                      <span className="w-1.5 h-1.5 bg-slate-400 rounded-full animate-bounce" style={{ animationDelay: "0ms" }}></span>
                      <span className="w-1.5 h-1.5 bg-slate-400 rounded-full animate-bounce" style={{ animationDelay: "150ms" }}></span>
                      <span className="w-1.5 h-1.5 bg-slate-400 rounded-full animate-bounce" style={{ animationDelay: "300ms" }}></span>
                    </div>
                  </div>
                )}
              </div>

              {/* Ô NHẬP TIN NHẮN CHAT */}
              <div className="flex gap-2">
                <input
                  type="text"
                  placeholder="Hỏi bất kỳ điều gì về Excel, Hàm thống kê, code Python, R..."
                  value={chatInput}
                  onChange={(e) => setChatInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") handleSendChatMessage();
                  }}
                  className="flex-1 px-4 py-2.5 bg-white border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500/20 transition font-sans shadow-sm"
                />
                <button
                  onClick={handleSendChatMessage}
                  disabled={chatLoading || !chatInput.trim()}
                  className="px-4.5 py-2.5 bg-indigo-600 hover:bg-indigo-500 disabled:bg-slate-100 disabled:text-slate-400 text-white text-xs font-bold rounded-xl transition flex items-center justify-center cursor-pointer shadow-lg shadow-indigo-100"
                >
                  <Send className="w-4 h-4" />
                </button>
              </div>

            </div>
          )}

        </div>

      </div>
      
    </div>
  );
}
