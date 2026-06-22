import React, { useState, useRef } from "react";
import { Document, Packer, Paragraph, TextRun, AlignmentType, Footer } from "docx";
import { FileText, FileDown, Settings, Copy, Check, Sparkles, AlertCircle, Loader2, RefreshCw, Upload, Languages, AlignLeft, CheckSquare, HelpCircle } from "lucide-react";
import { GoogleGenAI } from "@google/genai";

// === ĐỒNG BỘ CHUẨN ĐƯỜNG DẪN FILE VẬT LÝ VỚI BẢN 4.10.38 ===
import * as pdfjsLib from "pdfjs-dist";

// Sử dụng đuôi .mjs cho bản 4.10.38 để Vite 6 đóng gói trực tiếp thành URL
import pdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

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

  // Cấu hình định dạng Word
  const [docTitle, setDocTitle] = useState<string>("Tài liệu xuất bản");
  const [fontFamily, setFontFamily] = useState<string>("Calibri");
  const [fontSize, setFontSize] = useState<number>(12);
  const [lineSpacing, setLineSpacing] = useState<number>(1.15);
  const [addFooter, setAddFooter] = useState<boolean>(true);
  const [footerText, setFooterText] = useState<string>("Tài liệu được chuyển đổi từ Hệ thống VISC 2025");

  // Bộ AI tích hợp
  const [aiApiKey, setAiApiKey] = useState<string>(
    (import.meta as any).env?.VITE_GEMINI_API_KEY || ""
  );
  const [aiPrompt, setAiPrompt] = useState<string>("Hãy tóm tắt ngắn gọn các ý chính của tài liệu này một cách rành mạch dứt khoát:");
  const [aiResult, setAiResult] = useState<string>("");
  const [aiLoading, setAiLoading] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const dragRef = useRef<HTMLDivElement>(null);

  // Xử lý nạp File PDF
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processPdfFile(e.target.files[0]);
    }
  };

  const onDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (dragRef.current) {
      dragRef.current.classList.add("border-purple-500", "bg-purple-950/20");
    }
  };

  const onDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (dragRef.current) {
      dragRef.current.classList.remove("border-purple-500", "bg-purple-950/20");
    }
  };

  const onDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (dragRef.current) {
      dragRef.current.classList.remove("border-purple-500", "bg-purple-950/20");
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

  // Đọc nội dung file PDF chuyên sâu có căn lề bảo tồn xuống dòng
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
      
      const targetLib = getPdfjsLibrary();
      if (!targetLib) {
        throw new Error("Không thể kết nối tới cổng CDN của trình đọc thư viện PDF.");
      }

      const loadingTask = targetLib.getDocument({ data: arrayBuffer });
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

  // Gọi bộ tích hợp trí tuệ nhân tạo Google Gemini AI để xử lý văn bản trích xuất
  const handleRunAI = async () => {
    if (!extractedText) return;
    if (!aiApiKey) {
      setErrorMsg("Vui lòng bổ sung mã cấu hình API Key của Gemini để kích hoạt tính năng AI.");
      return;
    }

    setAiLoading(true);
    setAiResult("");
    try {
      // Khởi tạo SDK mới theo thư viện gốc @google/genai bạn đang dùng
      const ai = new GoogleGenAI({ apiKey: aiApiKey });
      const response = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: `${aiPrompt}\n\nNỘI DUNG TÀI LIỆU:\n${extractedText}`,
      });
      setAiResult(response.text || "Không có phản hồi dữ liệu từ mô hình AI.");
    } catch (err: any) {
      console.error(err);
      setErrorMsg("Quá trình kết nối AI thất bại: " + (err.message || "Lỗi không xác định"));
    } finally {
      setAiLoading(false);
    }
  };

  // Xử lý tạo và xuất File Word (.docx) hoàn chỉnh cấu trúc Section
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
                color: "111827"
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
                color: isPageHeader ? "8b5cf6" : "374151",
                bold: isPageHeader,
                italics: isPageHeader
              })
            ]
          })
        );
      });

      const sectionOptions: any = {
        properties: {
          page: {
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
          }
        },
        children: documentParagraphs
      };

      if (addFooter && footerText) {
        sectionOptions.footers = {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: footerText,
                    font: fontFamily,
                    size: 18,
                    color: "6b7280"
                  })
                ]
              })
            ]
          })
        };
      }

      const doc = new Document({
        sections: [sectionOptions]
      });

      const blob = await Packer.toBlob(doc);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${file ? file.name.replace(".pdf", "") : "tai-lieu"}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
      
      setLoadingStatus("Tải xuống tệp Word thành công!");
    } catch (err: any) {
      console.error(err);
      setErrorMsg("Không thể xuất file Word. Vui lòng thử lại.");
    } finally {
      setLoading(false);
    }
  };

  const copyToClipboard = () => {
    if (!extractedText) return;
