import React, { useState, useRef } from "react";
import * as pdfjsLib from "pdfjs-dist";
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  AlignmentType,
  Footer
} from "docx";
import { 
  FileText, 
  FileDown, 
  Settings, 
  Copy, 
  Check, 
  Sparkles, 
  AlertCircle, 
  Loader2, 
  RefreshCw, 
  Upload, 
  Languages, 
  AlignLeft, 
  CheckSquare,
  HelpCircle
} from "lucide-react";
import { GoogleGenAI } from "@google/genai";

const PDFJS_VERSION = pdfjsLib.version || "4.10.38";
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.worker.min.mjs`;

interface PageInfo {
  pageNumber: number;
  text: string;
}

const PdfToWord = React.memo(function PdfToWord() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [loadingStatus, setLoadingStatus] = useState<string>("");
  const [progress, setProgress] = useState<number>(0);
  const [extractedText, setExtractedText] = useState<string>("");
  const [pages, setPages] = useState<PageInfo[]>([]);
  const [copied, setCopied] = useState<boolean>(false);
  const [errorMsg, setErrorMsg] = useState<string>("");

  const [docTitle, setDocTitle] = useState<string>("Tài liệu xuất bản");
  const [fontFamily, setFontFamily] = useState<string>("Calibri");
  const [fontSize, setFontSize] = useState<number>(12);
  const [lineSpacing, setLineSpacing] = useState<number>(1.15);
  const [addFooter, setAddFooter] = useState<boolean>(true);
  const [footerText, setFooterText] = useState<string>("Tài liệu được chuyển đổi từ Hệ thống VISC");

  const [aiApiKey, setAiApiKey] = useState<string>(
    (import.meta as any).env?.VITE_GEMINI_API_KEY || ""
  );
  const [aiPrompt, setAiPrompt] = useState<string>("Hãy tóm tắt ngắn gọn các ý chính của tài liệu này một cách rành mạch dứt khoát:");
  const [aiResult, setAiResult] = useState<string>("");
  const [aiLoading, setAiLoading] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const dragRef = useRef<HTMLDivElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processPdfFile(e.target.files[0]);
    }
  };

  const onDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (dragRef.current) {
      dragRef.current.classList.add("border-indigo-500", "bg-indigo-50/50");
    }
  };

  const onDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (dragRef.current) {
      dragRef.current.classList.remove("border-indigo-500", "bg-indigo-50/50");
    }
  };

  const onDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (dragRef.current) {
      dragRef.current.classList.remove("border-indigo-500", "bg-indigo-50/50");
    }
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile.type === "application/pdf" || droppedFile.name.endsWith(".pdf")) {
        processPdfFile(droppedFile);
      } else {
        setErrorMsg("Hệ thống chỉ hỗ trợ xử lý tệp định dạng PDF (.pdf)");
      }
    }
  };

  const processPdfFile = async (pdfFile: File) => {
    setFile(pdfFile);
    setLoading(true);
    setErrorMsg("");
    setExtractedText("");
    setPages([]);
    setAiResult("");
    
    try {
      const arrayBuffer = await pdfFile.arrayBuffer();
      setLoadingStatus("Đang thiết lập cổng phân tích tài liệu...");
      
      const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
      const pdf = await loadingTask.promise;
      const numPages = pdf.numPages;
      const extractedPages: PageInfo[] = [];
      let fullText = "";

      for (let i = 1; i <= numPages; i++) {
        setLoadingStatus(`Đang đọc nội dung trang ${i} trên tổng số ${numPages}...`);
        setProgress(Math.round((i / numPages) * 100));
        
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        
        const items = textContent.items as any[];
        let pageText = "";
        let lastY = -1;
        
        for (const item of items) {
          if (item.str === undefined) continue;
          
          const currentY = item.transform ? item.transform[5] : -1;
          if (lastY !== -1 && Math.abs(currentY - lastY) > 5) {
            pageText += "\n";
          } else if (lastY !== -1) {
            pageText += " ";
          }
          
          pageText += item.str;
          lastY = currentY;
        }

        const trimmedPageText = pageText.trim();
        extractedPages.push({ pageNumber: i, text: trimmedPageText });
        fullText += `--- TRANG ${i} ---\n\n${trimmedPageText}\n\n`;
      }

      setPages(extractedPages);
      setExtractedText(fullText.trim());
      setLoadingStatus("Trích xuất nội dung thành công!");
    } catch (err: any) {
      console.error(err);
      setErrorMsg("Đã xảy ra lỗi khi giải nén tệp PDF. Vui lòng kiểm tra lại cấu trúc tệp của bạn.");
    } finally {
      setLoading(false);
      setProgress(0);
    }
  };

  const handleExportWord = async () => {
    if (!extractedText) return;
    
    setLoading(true);
    setLoadingStatus("Đang biên dịch tệp Word...");
    
    try {
      const documentParagraphs: Paragraph[] = [];

      if (docTitle) {
        documentParagraphs.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 240, after: 480 },
            children: [
              new TextRun({
                text: docTitle.toUpperCase(),
                bold: true,
                size: (fontSize + 6) * 2,
                font: fontFamily,
                color: "0f172a"
              })
            ]
          })
        );
      }

      const paragraphsList = extractedText.split("\n");
      
      paragraphsList.forEach((paraText) => {
        const text = paraText.trim();
        if (!text) {
          documentParagraphs.push(
            new Paragraph({
              spacing: { before: 120, after: 120 },
              children: [new TextRun({ text: "", font: fontFamily })]
            })
          );
          return;
        }

        const isPageHeader = text.startsWith("--- TRANG");
        
        documentParagraphs.push(
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { 
              before: 140, 
              after: 140, 
              line: lineSpacing * 240
            },
            children: [
              new TextRun({
                text: text,
                font: fontFamily,
                size: fontSize * 2,
                color: isPageHeader ? "4f46e5" : "1e293b",
                bold: isPageHeader,
                italics: isPageHeader
              })
            ]
          })
        );
      });

      const docOptions: any = {
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: 1440,
                  bottom: 1440,
                  left: 1440,
                  right: 1440
                }
              }
            },
            headers: {},
            footers: addFooter ? {
              default: new Footer({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: footerText,
                        font: fontFamily,
                        size: 9 * 2,
                        color: "64748b",
                        italics: true
                      })
                    ]
                  })
                ]
              })
            } : {},
            children: documentParagraphs
          }
        ]
      };

      const doc = new Document(docOptions);
      const blob = await Packer.toBlob(doc);
      
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${file ? file.name.replace(".pdf", "") : "Tai_lieu"}_VISC.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
      
      setLoadingStatus("Tải xuống thành công!");
    } catch (err) {
      console.error(err);
      setErrorMsg("Không thể thiết lập tệp Word do lỗi dịch định dạng.");
    } finally {
      setLoading(false);
    }
  };

  const handleCopyText = () => {
    if (!extractedText) return;
    navigator.clipboard.writeText(extractedText);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleAiRefinement = async () => {
    if (!extractedText) return;
    if (!aiApiKey) {
      setErrorMsg("Cần cấu hình khoá API VITE_GEMINI_API_KEY để sử dụng trợ lý AI.");
      return;
    }

    setAiLoading(true);
    setAiResult("");
    
    try {
      const ai = new GoogleGenAI({
        apiKey: aiApiKey,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build"
          }
        }
      });

      const previewText = extractedText.substring(0, 15000);
      const systemPrompt = "Bạn là trợ lý phân tích văn bản chuyên nghiệp của Hệ thống so sánh tổng hợp VISC. Hãy làm việc hoàn hảo và đưa ra kết quả ngắn gọn nhất bằng tiếng Việt.";
      
      const response = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: `${aiPrompt}\n\n[ĐOẠN TRÍCH TÀI LIỆU PDF]:\n${previewText}`,
        config: {
          systemInstruction: systemPrompt,
          temperature: 0.3
        }
      });

      if (response && response.text) {
        setAiResult(response.text);
      } else {
        setAiResult("Mô hình phản hồi không hợp lệ.");
      }
    } catch (err: any) {
      console.error(err);
      setAiResult(`Lỗi AI: ${err.message || "Không thể tải phản hồi từ mô hình. Hãy kiểm tra lại khoá API."}`);
    } finally {
      setAiLoading(false);
    }
  };

  const handleReset = () => {
    setFile(null);
    setExtractedText("");
    setPages([]);
    setAiResult("");
    setErrorMsg("");
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  return (
    <div className="space-y-6 font-sans text-slate-850">
      <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-xl">
        
        {/* Tiêu đề & Giới thiệu */}
        <div className="border-b border-slate-100 pb-4 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div>
            <h3 className="text-xl font-black text-[#0f172a] flex items-center gap-2.5">
              <FileText className="w-5.5 h-5.5 text-indigo-600 animate-pulse" />
              ĐỌC FILE PDF & CHUYỂN SANG WORD (.DOCX) CHUẨN CAO
            </h3>
            <p className="text-xs text-slate-500 mt-1">
              Giải pháp trích xuất trực tiếp văn bản từ file PDF gốc hoàn toàn bằng cơ chế trình duyệt an toàn bảo mật, tùy chỉnh cấu hình lề và tải về Word sạch sẽ.
            </p>
          </div>
          
          {file && (
            <button
              onClick={handleReset}
              className="text-xs text-rose-700 hover:text-white hover:bg-rose-700 border border-rose-250 bg-rose-50 px-3.5 py-1.5 rounded-lg transition-all cursor-pointer flex items-center gap-1.5 font-bold"
            >
              <RefreshCw className="w-3.5 h-3.5" /> Nạp tệp khác
            </button>
          )}
        </div>

        {errorMsg && (
          <div className="bg-rose-50 border border-rose-200 rounded-xl p-4 flex items-start gap-3">
            <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
            <p className="text-xs text-rose-800 leading-relaxed font-semibold">{errorMsg}</p>
          </div>
        )}

        {/* Khung tải tệp PDF lên nếu chưa có file */}
        {!file && (
          <div
            ref={dragRef}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
            onClick={() => fileInputRef.current?.click()}
            className="border-2 border-dashed border-slate-300 hover:border-indigo-500 rounded-2xl p-10 text-center transition-all cursor-pointer bg-slate-50 hover:bg-indigo-50/20 group space-y-4"
            id="pdf-upload-dropzone"
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileChange}
              accept=".pdf"
              className="hidden"
            />
            <div className="bg-white w-14 h-14 rounded-full flex items-center justify-center mx-auto border border-slate-200 group-hover:scale-105 transition-transform shadow-sm">
              <Upload className="w-6 h-6 text-indigo-650" />
            </div>
            <div className="space-y-1.5">
              <h4 className="text-sm font-bold text-slate-800">
                Kéo thả file PDF của bạn vào đây hoặc <span className="text-indigo-600 underline">nhấp chuột để chọn tệp</span>
              </h4>
              <p className="text-xs text-slate-500">
                Chỉ hỗ trợ tệp tài liệu PDF (.pdf), kích thước tối đa gợi ý 50MB.
              </p>
            </div>
          </div>
        )}

        {/* Trạng thái tiến trình giải nén */}
        {loading && (
          <div className="bg-slate-50 border border-slate-200 rounded-2xl p-6 flex flex-col items-center justify-center text-center space-y-4">
            <Loader2 className="w-8 h-8 text-indigo-600 animate-spin" />
            <div className="space-y-1.5 w-full max-w-xs">
              <div className="text-xs text-slate-700 font-bold">{loadingStatus}</div>
              {progress > 0 && (
                <div className="w-full bg-slate-200 rounded-full h-1.5">
                  <div 
                    className="bg-indigo-600 h-1.5 rounded-full transition-all duration-300"
                    style={{ width: `${progress}%` }}
                  ></div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Kết quả sau khi trích xuất */}
        {extractedText && !loading && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            
            {/* Cột trái: Văn bản đã trích xuất & Xem trước */}
            <div className="lg:col-span-7 space-y-4 flex flex-col">
              <div className="flex justify-between items-center">
                <span className="text-xs font-bold text-indigo-800 uppercase tracking-wider font-mono">
                  📝 Văn bản trích xuất ({pages.length} trang, {extractedText.length.toLocaleString()} ký tự)
                </span>
                
                <button
                  onClick={handleCopyText}
                  className="bg-white hover:bg-slate-50 text-xs px-3.5 py-1.5 rounded-xl border border-slate-300 hover:border-slate-400 cursor-pointer flex items-center gap-1.5 transition-all text-slate-800 font-bold shadow-sm"
                >
                  {copied ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-600" />
                      <span className="text-emerald-700 font-bold">Đã sao chép!</span>
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5 text-slate-500" />
                      <span>Sao chép tất cả</span>
                    </>
                  )}
                </button>
              </div>

              {/* Ô hiển thị văn bản và hỗ trợ chỉnh sửa trực tiếp */}
              <div className="relative border border-slate-300 rounded-xl overflow-hidden bg-slate-50 flex-1">
                <textarea
                  value={extractedText}
                  onChange={(e) => setExtractedText(e.target.value)}
                  className="w-full h-[400px] bg-transparent text-slate-800 p-4 font-mono text-xs leading-relaxed border-0 focus:ring-0 resize-none overflow-y-auto"
                ></textarea>
                <div className="absolute bottom-2.5 right-3 text-[9px] text-slate-500 font-bold bg-white/90 px-2 py-1 rounded border border-slate-200">
                  Bạn có thể điền, sửa hoặc tinh chỉnh trực tiếp văn bản này trước khi tải Word
                </div>
              </div>
            </div>

            {/* Cột phải: Cấu hình tải file .docx & Trợ lý AI */}
            <div className="lg:col-span-5 space-y-6">
              
              {/* Thẻ cấu hình xuất Word */}
              <div className="bg-slate-50 border border-slate-200 rounded-2xl p-5 space-y-4">
                <h4 className="text-sm font-bold text-[#0f172a] flex items-center gap-2">
                  <Settings className="w-4.5 h-4.5 text-indigo-600" />
                  CẤU HÌNH BIÊN DỊCH TỆP WORD (.DOCX)
                </h4>
                
                <div className="space-y-3.5">
                  
                  {/* Tiêu đề tài liệu */}
                  <div>
                    <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest block mb-1">
                      🏷️ Tiêu đề tài liệu (In trang đầu)
                    </label>
                    <input
                      type="text"
                      value={docTitle}
                      onChange={(e) => setDocTitle(e.target.value)}
                      className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium font-sans"
                      placeholder="Nhập tiêu đề hoặc để trống..."
                    />
                  </div>

                  {/* Font chữ & Cỡ chữ */}
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest block mb-1">
                        🔤 Font chữ
                      </label>
                      <select
                        value={fontFamily}
                        onChange={(e) => setFontFamily(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 cursor-pointer font-medium"
                      >
                        <option value="Calibri">Calibri (Mặc định)</option>
                        <option value="Times New Roman">Times New Roman</option>
                        <option value="Arial">Arial</option>
                        <option value="Segoe UI">Segoe UI</option>
                      </select>
                    </div>

                    <div>
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest block mb-1">
                        📏 Cỡ chữ cơ bản
                      </label>
                      <select
                        value={fontSize}
                        onChange={(e) => setFontSize(parseInt(e.target.value))}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 cursor-pointer font-medium"
                      >
                        <option value={11}>11pt (Nhỏ gọn)</option>
                        <option value={12}>12pt (Tiêu chuẩn)</option>
                        <option value={13}>13pt (Dễ đọc)</option>
                        <option value={14}>14pt (Dành cho báo cáo)</option>
                      </select>
                    </div>
                  </div>

                  {/* Khoảng cách dòng */}
                  <div>
                    <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest block mb-1">
                      ↕️ Khoảng cách đoạn
                    </label>
                    <select
                      value={lineSpacing}
                      onChange={(e) => setLineSpacing(parseFloat(e.target.value))}
                      className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 cursor-pointer font-medium"
                    >
                      <option value={1.0}>Đơn (1.0)</option>
                      <option value={1.15}>Tiêu chuẩn (1.15)</option>
                      <option value={1.5}>Vừa phải (1.5)</option>
                    </select>
                  </div>

                  {/* Chân trang Footer tùy biến */}
                  <div className="pt-2 border-t border-slate-200 space-y-2">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={addFooter}
                        onChange={(e) => setAddFooter(e.target.checked)}
                        className="rounded border-slate-300 text-indigo-650 focus:ring-indigo-500 bg-white"
                      />
                      <span className="text-xs text-slate-700 font-bold">Bật footer thông tin chân trang</span>
                    </label>

                    {addFooter && (
                      <input
                        type="text"
                        value={footerText}
                        onChange={(e) => setFooterText(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-[11px] text-slate-650 focus:ring-1 focus:ring-indigo-500 font-medium"
                      />
                    )}
                  </div>

                  {/* Bấm tải tệp Word */}
                  <button
                    onClick={handleExportWord}
                    className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs py-3 rounded-xl transition-all cursor-pointer flex items-center justify-center gap-1.5 shadow-md shadow-indigo-100"
                  >
                    <FileDown className="w-4 h-4 text-indigo-200" /> BIÊN DỊCH VÀ XUẤT FILE WORD (.DOCX)
                  </button>
                </div>
              </div>

              {/* Thẻ tích hợp AI */}
              <div className="bg-slate-50 border border-slate-200 rounded-2xl p-5 space-y-4">
                <div className="flex items-center justify-between">
                  <h4 className="text-sm font-bold text-[#0f172a] flex items-center gap-2">
                    <Sparkles className="w-4.5 h-4.5 text-indigo-600 animate-pulse" />
                    TRỢ LÝ TRÍ TUỆ NHÂN TẠO AI
                  </h4>
                  <span className="text-[9px] bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded font-mono font-bold uppercase tracking-wider border border-indigo-200">
                    Gemini Active
                  </span>
                </div>

                <div className="space-y-3">
                  <div>
                    <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest block mb-1">
                      ⚙️ Chọn mẫu yêu cầu AI:
                    </label>
                    <select
                      onChange={(e) => setAiPrompt(e.target.value)}
                      className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 cursor-pointer font-semibold text-indigo-900"
                    >
                      <option value="Hãy tóm tắt ngắn gọn các ý chính của tài liệu này một cách rành mạch dứt khoát:">📝 Tóm tắt ý chính cốt lõi tài liệu</option>
                      <option value="Hãy dịch văn bản này sang dạng tiếng Việt hoàn toàn chuẩn hóa, sửa các viết tắt và các sai sót dịch thuật (nếu là tài liệu gốc ngoại ngữ), hoặc dịch từ Việt sang Anh ngữ:">🌐 Dịch thuật song ngữ Việt - Anh</option>
                      <option value="Phát hiện và chỉnh sửa lại toàn bộ các ký tự lỗi OCR chính tả do scan hoặc trích xuất không chuẩn, trình bày lại dưới dạng các đoạn phân mục sạch sẽ, giữ nguyên nghĩa:">🧼 Sạch hoá lỗi OCR & Sắp xếp lại chữ</option>
                      <option value="Trích xuất các danh mục thông tin gồm Mã số thuế, Tên doanh nghiệp, Ngành hoạt động, Người đại diện hoặc địa điểm xuất hiện trong văn bản này xuất ra một cấu trúc thống kê đẹp mắt:">📊 Trích xuất danh sách doanh nghiệp ẩn</option>
                    </select>
                  </div>

                  <div>
                    <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest block mb-1">
                      ✍️ Tự viết Prompt yêu cầu AI:
                    </label>
                    <textarea
                      value={aiPrompt}
                      onChange={(e) => setAiPrompt(e.target.value)}
                      rows={2}
                      className="w-full bg-white border border-slate-300 rounded-lg p-2 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-sans font-medium"
                    />
                  </div>

                  {!aiApiKey && (
                    <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-[10px] text-amber-800 flex items-start gap-2 leading-relaxed">
                      <HelpCircle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                      <div>
                        Để khởi chạy tác vụ AI, bạn hãy khai báo khóa <strong>VITE_GEMINI_API_KEY</strong> trong phần <strong>Cài đặt &gt; Khóa bí mật (Secrets)</strong> của AI Studio.
                      </div>
                    </div>
                  )}

                  <button
                    onClick={handleAiRefinement}
                    disabled={aiLoading || !extractedText || !aiApiKey}
                    className={`w-full text-xs font-bold py-2.5 rounded-xl transition-all flex items-center justify-center gap-1.5 ${
                      !extractedText || !aiApiKey
                        ? "bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed"
                        : "bg-indigo-600 hover:bg-indigo-700 text-white cursor-pointer shadow-md hover:shadow-indigo-100"
                    }`}
                  >
                    {aiLoading ? (
                      <>
                        <Loader2 className="w-4 h-4 text-white animate-spin" />
                        ĐANG PHÂN TÍCH &amp; SUY NGHĨ...
                      </>
                    ) : (
                      <>
                        <Sparkles className="w-4 h-4 text-indigo-200" />
                        GỬI YÊU CẦU CHO TRỢ LÝ GEMINI
                      </>
                    )}
                  </button>

                  {/* Kết quả phản hồi AI */}
                  {aiResult && (
                    <div className="mt-3 bg-white border border-slate-200 rounded-xl p-4 space-y-3.5 shadow-sm">
                      <div className="text-[11px] font-bold text-indigo-700 uppercase tracking-wide font-mono flex items-center gap-1.5 border-b border-slate-150 pb-1.5">
                        <CheckSquare className="w-3.5 h-3.5 text-indigo-600" />
                        PHẢN HỒI PHÂN TÍCH TỪ AI
                      </div>
                      <div className="text-xs text-slate-800 font-medium leading-relaxed whitespace-pre-wrap max-h-[220px] overflow-y-auto pr-1">
                        {aiResult}
                      </div>

                      <div className="pt-2 border-t border-slate-200 flex items-center gap-2">
                        <button
                          onClick={() => {
                            navigator.clipboard.writeText(aiResult);
                            alert("Đã sao chép phản hồi AI vào khay nhớ tạm!");
                          }}
                          className="text-[10px] bg-slate-50 hover:bg-slate-100 text-slate-700 font-bold px-2.5 py-1.5 rounded-lg border border-slate-300 transition-colors cursor-pointer"
                        >
                          Sao chép phản hồi AI
                        </button>
                        <button
                          onClick={() => {
                            setExtractedText((prev) => `${prev}\n\n--- PHẢN HỒI PHÂN TÍCH TỪ AI ---\n\n${aiResult}`);
                          }}
                          className="text-[10px] bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold px-2.5 py-1.5 rounded-lg border border-indigo-200 transition-colors cursor-pointer"
                        >
                          Ghép vào tài liệu để xuất
                        </button>
                      </div>
                    </div>
                  )}

                </div>
              </div>

            </div>
          </div>
        )}

      </div>
    </div>
  );
});

PdfToWord.displayName = "PdfToWord";

export default PdfToWord;
