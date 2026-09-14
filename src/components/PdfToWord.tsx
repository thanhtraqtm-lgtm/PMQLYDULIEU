import React, { useState } from "react";
import { FileText, Upload, Download, Copy, Check, RefreshCw, AlertCircle } from "lucide-react";

export const PdfToWord: React.FC = () => {
  const [extractedText, setExtractedText] = useState<string>("");
  const [fileName, setFileName] = useState<string>("");
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [copied, setCopied] = useState<boolean>(false);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setIsProcessing(true);

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        // Simple extraction fallback using text decoder for raw text streams
        const decoder = new TextDecoder("utf-8");
        const raw = decoder.decode(buffer);
        // Clean ASCII/UTF-8 streams from PDF
        const textParts: string[] = [];
        const matches = raw.match(/\(([^)]+)\)\s*Tj/g) || raw.match(/\[([^\]]+)\]\s*TJ/g);
        if (matches && matches.length > 0) {
          matches.forEach(m => {
            const clean = m.replace(/[\(\)\[\]]/g, "").replace(/Tj|TJ/g, "").trim();
            if (clean) textParts.push(clean);
          });
          setExtractedText(textParts.join(" "));
        } else {
          setExtractedText(`[Nội dung tệp PDF: ${file.name}]\nKích thước: ${(file.size / 1024).toFixed(1)} KB\nTệp đã sẵn sàng để trích xuất văn bản hoặc chuyển đổi.`);
        }
      } catch (err: any) {
        setExtractedText(`Không thể đọc trực tiếp định dạng tệp: ${err.message}`);
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const handleDownloadDoc = () => {
    if (!extractedText) return;
    // Download as HTML document readable by Microsoft Word (.doc)
    const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' "+
                   "xmlns:w='urn:schemas-microsoft-com:office:word' "+
                   "xmlns='http://www.w3.org/TR/REC-html40'>"+
                   "<head><meta charset='utf-8'><title>Export Word</title></head><body>";
    const footer = "</body></html>";
    const sourceHTML = header + `<div style="font-family: Arial, sans-serif; font-size: 14pt; line-height: 1.6;">${extractedText.replace(/\n/g, "<br/>")}</div>` + footer;

    const source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
    const fileDownload = document.createElement("a");
    document.body.appendChild(fileDownload);
    fileDownload.href = source;
    fileDownload.download = `${fileName.replace(/\.[^/.]+$/, "") || "ChuyenDoi"}.doc`;
    fileDownload.click();
    document.body.removeChild(fileDownload);
  };

  const handleCopy = () => {
    if (!extractedText) return;
    navigator.clipboard.writeText(extractedText);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 space-y-6">
      <div className="flex flex-wrap items-center justify-between gap-4 pb-4 border-b border-slate-100">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-red-100 flex items-center justify-center text-red-600">
            <FileText className="w-5 h-5" />
          </div>
          <div>
            <h3 className="font-bold text-slate-800 text-base">Chuyển Đổi PDF Sang Văn Bản Word (.doc)</h3>
            <p className="text-xs text-slate-500">
              Trích xuất nội dung từ tài liệu biểu mẫu báo cáo PDF sang định dạng Microsoft Word dễ dàng chỉnh sửa.
            </p>
          </div>
        </div>

        {extractedText && (
          <div className="flex items-center gap-2">
            <button
              onClick={handleCopy}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg border border-slate-300 transition-colors cursor-pointer"
            >
              {copied ? <Check className="w-4 h-4 text-emerald-600" /> : <Copy className="w-4 h-4" />}
              {copied ? "Đã sao chép!" : "Sao Chép"}
            </button>
            <button
              onClick={handleDownloadDoc}
              className="inline-flex items-center gap-1.5 px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg shadow-sm transition-colors cursor-pointer"
            >
              <Download className="w-4 h-4" />
              Tải File Word (.doc)
            </button>
          </div>
        )}
      </div>

      {/* Upload zone */}
      <div className="border-2 border-dashed border-slate-300 rounded-xl p-8 text-center bg-slate-50/50 hover:bg-slate-50 transition-colors">
        <Upload className="w-8 h-8 text-slate-400 mx-auto mb-3" />
        <p className="text-sm font-bold text-slate-700">Kéo thả tệp PDF vào đây hoặc duyệt tệp từ máy tính</p>
        <p className="text-xs text-slate-400 mt-1 mb-4">Hỗ trợ các văn bản báo cáo, nghị quyết, thông tư thống kê dạng PDF</p>
        <label className="inline-flex items-center gap-2 px-4 py-2 bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold rounded-lg shadow-sm cursor-pointer transition-colors">
          <Upload className="w-4 h-4" />
          Chọn Tệp PDF
          <input type="file" accept=".pdf" onChange={handleFileUpload} className="hidden" />
        </label>
      </div>

      {/* Preview */}
      {isProcessing && (
        <div className="py-8 text-center text-slate-500 text-xs flex items-center justify-center gap-2">
          <RefreshCw className="w-4 h-4 animate-spin text-blue-600" />
          Đang phân tích cấu trúc tệp PDF...
        </div>
      )}

      {extractedText && (
        <div className="space-y-2">
          <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider">
            Nội Dung Văn Bản Trích Xuất Được ({fileName}):
          </label>
          <textarea
            rows={12}
            value={extractedText}
            onChange={(e) => setExtractedText(e.target.value)}
            className="w-full text-xs font-mono p-4 border border-slate-300 rounded-xl bg-slate-50/30 focus:outline-none focus:ring-2 focus:ring-blue-500/20"
          />
        </div>
      )}
    </div>
  );
};
export default PdfToWord;
