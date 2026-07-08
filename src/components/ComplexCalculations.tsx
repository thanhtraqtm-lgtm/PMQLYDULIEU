import React, { useState, useEffect, useRef } from "react";
import { Sliders, Database, Play, CheckCircle, AlertTriangle, FileSpreadsheet, Sparkles, RefreshCw, ChevronRight, CornerDownRight, HelpCircle } from "lucide-react";
import { parseRobustNumber } from "./sectorRevenueChart";

interface ComplexCalculationsProps {
  mathFileAId: string;
  setMathFileAId: (val: string) => void;
  mathFileBId: string;
  setMathFileBId: (val: string) => void;
  mathKeyA: string;
  setMathKeyA: (val: string) => void;
  mathKeyA2: string;
  setMathKeyA2: (val: string) => void;
  mathKeyB: string;
  setMathKeyB: (val: string) => void;
  mathKeyB2: string;
  setMathKeyB2: (val: string) => void;
  mathColA: string;
  setMathColA: (val: string) => void;
  mathColA2: string;
  setMathColA2: (val: string) => void;
  mathColA3: string;
  setMathColA3: (val: string) => void;
  mathColB: string;
  setMathColB: (val: string) => void;
  mathColB2: string;
  setMathColB2: (val: string) => void;
  mathColB3: string;
  setMathColB3: (val: string) => void;
  mathOp: "+" | "-" | "*" | "/";
  setMathOp: (val: "+" | "-" | "*" | "/") => void;
  mathOp2: "+" | "-" | "*" | "/";
  setMathOp2: (val: "+" | "-" | "*" | "/") => void;
  mathOp3: "+" | "-" | "*" | "/";
  setMathOp3: (val: "+" | "-" | "*" | "/") => void;
  mathNewColName: string;
  setMathNewColName: (val: string) => void;
  mathNewColName2: string;
  setMathNewColName2: (val: string) => void;
  mathNewColName3: string;
  setMathNewColName3: (val: string) => void;
  mathFilterA: string;
  setMathFilterA: (val: string) => void;
  mathFilterB: string;
  setMathFilterB: (val: string) => void;
  allAvailableFiles: Array<{
    id: string;
    name: string;
    data: any[];
    columns: string[];
  }>;
  handlePerformCrossFileMath: () => void;
}

interface LogEntry {
  time: string;
  type: "system" | "success" | "error" | "warning" | "info";
  message: string;
}

export default function ComplexCalculations({
  mathFileAId,
  setMathFileAId,
  mathFileBId,
  setMathFileBId,
  mathKeyA,
  setMathKeyA,
  mathKeyA2,
  setMathKeyA2,
  mathKeyB,
  setMathKeyB,
  mathKeyB2,
  setMathKeyB2,
  mathColA,
  setMathColA,
  mathColA2,
  setMathColA2,
  mathColA3,
  setMathColA3,
  mathColB,
  setMathColB,
  mathColB2,
  setMathColB2,
  mathColB3,
  setMathColB3,
  mathOp,
  setMathOp,
  mathOp2,
  setMathOp2,
  mathOp3,
  setMathOp3,
  mathNewColName,
  setMathNewColName,
  mathNewColName2,
  setMathNewColName2,
  mathNewColName3,
  setMathNewColName3,
  mathFilterA,
  setMathFilterA,
  mathFilterB,
  setMathFilterB,
  allAvailableFiles,
  handlePerformCrossFileMath
}: ComplexCalculationsProps) {
  
  const [logs, setLogs] = useState<LogEntry[]>([
    { time: getFormattedTime(), type: "system", message: "Đang gọi bộ kết nối dữ liệu ngầm từ hệ thống..." },
    { time: getFormattedTime(), type: "system", message: "Hệ thống kết nối cơ sở dữ liệu đa tầng khởi động thành công!" },
    { time: getFormattedTime(), type: "info", message: "Sẵn sàng nạp tệp ma trận đồng bộ VSIC. Vui lòng thiết lập cấu hình." }
  ]);
  const logEndRef = useRef<HTMLDivElement>(null);

  function getFormattedTime() {
    const d = new Date();
    return d.toTimeString().split(" ")[0];
  }

  function addLog(message: string, type: "system" | "success" | "error" | "warning" | "info" = "info") {
    setLogs(prev => [...prev, { time: getFormattedTime(), type, message }]);
  }

  useEffect(() => {
    if (logEndRef.current) {
      logEndRef.current.scrollIntoView({ behavior: "smooth" });
    }
  }, [logs]);

  useEffect(() => {
    if (mathFileAId) {
      const file = allAvailableFiles.find(f => f.id === mathFileAId);
      if (file) {
        addLog(`[TỆP A] Đã kết nối thành công: "${file.name}" (${file.data.length} dòng, ${file.columns.length} cột).`, "success");
      }
    }
  }, [mathFileAId, allAvailableFiles]);

  useEffect(() => {
    if (mathFileBId) {
      const file = allAvailableFiles.find(f => f.id === mathFileBId);
      if (file) {
        addLog(`[TỆP B] Đã kết nối thành công: "${file.name}" (${file.data.length} dòng, ${file.columns.length} cột).`, "success");
      }
    }
  }, [mathFileBId, allAvailableFiles]);

  useEffect(() => {
    if (mathKeyA && mathKeyB) {
      addLog(`[ĐỐI CHIẾU] Đã ánh xạ liên kết dòng: Tệp A [${mathKeyA}] <==> Tệp B [${mathKeyB}].`, "info");
    }
  }, [mathKeyA, mathKeyB]);

  const handleAutoConnect = () => {
    const fileA = allAvailableFiles.find(f => f.id === mathFileAId);
    const fileB = allAvailableFiles.find(f => f.id === mathFileBId);
    
    addLog("Đang kích hoạt bộ giải toán tự động kết nối ma trận đồng bộ...", "system");

    if (fileA && fileB) {
      const commonKeys = fileA.columns.filter(c => fileB.columns.includes(c));
      const possibleKey = commonKeys.find(c => 
        c.toLowerCase().includes("ngành") || 
        c.toLowerCase().includes("tên") || 
        c.toLowerCase().includes("mã") ||
        c.toLowerCase().includes("địa bàn") ||
        c.toLowerCase().includes("xã")
      ) || commonKeys[0];
      
      if (possibleKey) {
        setMathKeyA(possibleKey);
        setMathKeyB(possibleKey);
        addLog(`[AUTO-LINK] Đã tự động chọn cột khóa chung: "${possibleKey}"`, "success");
      } else {
        addLog("[AUTO-LINK] Cảnh báo: Không tìm thấy cột khóa chung rõ ràng. Bạn cần chọn thủ công.", "warning");
      }

      const possibleNumCols = fileA.columns.filter(c => 
        !c.toLowerCase().includes("ngành") && 
        !c.toLowerCase().includes("tên") && 
        !c.toLowerCase().includes("mã") &&
        !c.toLowerCase().includes("địa bàn") &&
        !c.toLowerCase().includes("xã") &&
        fileB.columns.includes(c)
      );

      if (possibleNumCols.length > 0) {
        setMathColA(possibleNumCols[0]);
        setMathColB(possibleNumCols[0]);
        addLog(`[AUTO-LINK] Khớp cột số cặp 1: A [${possibleNumCols[0]}] & B [${possibleNumCols[0]}]`, "info");
        
        if (possibleNumCols.length > 1) {
          setMathColA2(possibleNumCols[1]);
          setMathColB2(possibleNumCols[1]);
          addLog(`[AUTO-LINK] Khớp cột số cặp 2: A [${possibleNumCols[1]}] & B [${possibleNumCols[1]}]`, "info");
        }
        if (possibleNumCols.length > 2) {
          setMathColA3(possibleNumCols[2]);
          setMathColB3(possibleNumCols[2]);
          addLog(`[AUTO-LINK] Khớp cột số cặp 3: A [${possibleNumCols[2]}] & B [${possibleNumCols[2]}]`, "info");
        }
        
        addLog("Đã kết nối các cột trùng tên và các chỉ tiêu số đồng bộ thành công!", "success");
      } else {
        addLog("[AUTO-LINK] Không tìm thấy cột số có tên giống nhau. Hãy chỉ định cột phép tính bên dưới.", "warning");
      }
    } else {
      addLog("Thất bại: Vui lòng nạp và chọn đầy đủ cả Tệp A và Tệp B để sử dụng tính năng kết nối tự động!", "error");
    }
  };

  const handleExecuteWithLogs = () => {
    if (!mathFileAId || !mathFileBId) {
      addLog("Lỗi thao tác: Chưa chọn đầy đủ Tệp A và Tệp B để làm phép toán!", "error");
      addLog("Hãy chắc chắn bạn đã nạp các tệp dữ liệu nông nghiệp, dân số hoặc công nghiệp.", "warning");
      alert("Vui lòng chọn đầy đủ cả Tệp A và Tệp B!");
      return;
    }
    if (!mathKeyA || !mathKeyB) {
      addLog("Lỗi thao tác: Chưa cấu hình cột khóa liên kết dòng để đồng bộ ma trận!", "error");
      alert("Vui lòng chọn cột khóa để liên kết dòng giữa 2 tệp!");
      return;
    }
    if (!mathColA || !mathColB) {
      addLog("Lỗi thao tác: Cột số so sánh 1 không được để trống!", "error");
      alert("Vui lòng chọn cột số ở cả 2 tệp cho Cặp 1!");
      return;
    }

    addLog("Đang biên dịch thuật toán gộp nhóm & giải ma trận dữ liệu...", "system");
    addLog(`Đang thực hiện phép toán: [${mathColA}] ${mathOp} [${mathColB}]...`, "info");
    
    try {
      handlePerformCrossFileMath();
      addLog(`Thực hiện phép toán thành công! Đã gộp và tạo tệp kết quả so sánh mới.`, "success");
      addLog(`Cột kết quả mới: "${mathNewColName}" đã được chèn vào ma trận.`, "success");
    } catch (e: any) {
      addLog(`Lỗi xử lý ma trận: ${e.message || e}`, "error");
    }
  };

  return (
    <div className="bg-slate-50 p-4 sm:p-6 rounded-2xl border border-slate-200 space-y-6 text-slate-800 shadow-xl font-sans">
      
      {/* KHU VỰC BANNER HƯỚNG DẪN HOẠT ĐỘNG */}
      <div className="bg-gradient-to-r from-indigo-900 to-slate-900 rounded-2xl p-5 text-white space-y-3 shadow-md border border-indigo-950 relative overflow-hidden">
        <div className="absolute -right-10 -top-10 opacity-10 pointer-events-none">
          <Sparkles className="w-40 h-40 text-white" />
        </div>
        <div className="flex items-center gap-2">
          <span className="bg-amber-400 text-indigo-950 font-black text-[10px] uppercase tracking-wider px-2 py-0.5 rounded shadow-sm">
            HƯỚNG DẪN NHANH
          </span>
          <h4 className="text-sm font-bold text-slate-100 flex items-center gap-1.5">
            💡 Cách sử dụng Phép tính Liên kết Chéo giữa các tệp dữ liệu
          </h4>
        </div>
        <p className="text-xs text-indigo-200 leading-relaxed max-w-4xl font-medium font-sans">
          Tính năng này cho phép bạn làm phép toán <strong className="text-amber-300 font-sans">Cộng (+), Trừ (-), Nhân (*), Chia (/)</strong> trực tiếp giữa các cột số của hai tệp tin khác nhau (ví dụ: đối chiếu số liệu năm nay với năm cũ, đối sánh các tệp nông nghiệp và công nghiệp...).
        </p>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3 pt-1 text-[11px] text-slate-300 font-sans">
          <div className="bg-slate-800/40 p-3 rounded-lg border border-slate-700/50">
            <span className="font-bold text-amber-300 block mb-1">1. CHỌN TỆP &amp; KHÓA (BƯỚC 1)</span>
            Chọn Tệp A và Tệp B, sau đó chọn <strong className="font-semibold text-white">Cột khóa chung</strong> (Ví dụ: tên Xã hoặc mã Ngành). Hệ thống sẽ liên kết dòng tương ứng trong thời gian thực <em className="text-indigo-200">(không cần nút bấm nạp vì dữ liệu cập nhật tự động)</em>.
          </div>
          <div className="bg-slate-800/40 p-3 rounded-lg border border-slate-700/50">
            <span className="font-bold text-amber-300 block mb-1">2. THIẾT LẬP PHÉP TÍNH (BƯỚC 2)</span>
            Chọn các cột số của Tệp A, phép tính mong muốn, cột tương ứng của Tệp B và đặt tên cho cột kết quả mới sẽ tạo ra.
          </div>
          <div className="bg-slate-800/40 p-3 rounded-lg border border-slate-700/50">
            <span className="font-bold text-amber-300 block mb-1">3. CHẠY KẾT QUẢ (NÚT BẤM)</span>
            Cuộn xuống cuối Bước 2 và bấm nút <strong className="text-emerald-400 font-semibold">"CHẠY PHÉP TÍNH LIÊN KẾT CHÉO"</strong> màu xanh để tạo tệp tổng hợp so sánh gộp mới!
          </div>
        </div>
      </div>

      {/* 1. CẤU HÌNH CÁC TỆP TIN VÀ KHÓA LIÊN KẾT ĐỒNG BỘ */}
      <fieldset className="border border-slate-200 rounded-xl p-4 sm:p-5 bg-white text-slate-800 space-y-4 shadow-sm relative">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-2 mb-2">
          <span className="text-xs sm:text-[12.5px] font-bold text-indigo-900 uppercase tracking-wide">
            1. CHỌN TỆP PHÂN TÍCH &amp; THIẾT LẬP KHÓA ÁNH XẠ
          </span>
          <button
            onClick={handleAutoConnect}
            className="bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-300 rounded-lg px-2.5 py-1 text-xs font-bold transition-all flex items-center gap-1 active:scale-95 cursor-pointer self-start sm:self-auto"
            title="Tự động tìm cột khóa chung và khớp các cột chỉ tiêu tương đồng giữa 2 tệp"
          >
            <Sparkles className="w-3.5 h-3.5 text-amber-600" /> Tự động khớp cột nhanh
          </button>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Tệp nguồn A */}
          <div className="space-y-3 p-4 bg-slate-50 rounded-xl border border-slate-200">
            <span className="text-[11px] font-bold text-indigo-900 block uppercase tracking-wider flex items-center gap-1.5 border-b border-slate-200 pb-1.5">
              <span className="w-2.5 h-2.5 rounded-full bg-indigo-600 block"></span> TỆP TIN NGUỒN A
            </span>
            
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-slate-500 block uppercase">Chọn tệp tin A:</label>
                <select
                  value={mathFileAId}
                  onChange={(e) => {
                    setMathFileAId(e.target.value);
                    setMathColA("");
                    setMathColA2("");
                    setMathColA3("");
                    setMathKeyA("");
                    setMathKeyA2("");
                  }}
                  className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                >
                  <option value="">-- Chọn tệp A --</option>
                  {allAvailableFiles.map(f => (
                    <option key={f.id} value={f.id}>{f.name.replace(/📂\s*\[TỆP CHÍNH\]\s*/, "")}</option>
                  ))}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-slate-500 block uppercase">Cột khóa chính (A):</label>
                <select
                  value={mathKeyA}
                  onChange={(e) => setMathKeyA(e.target.value)}
                  className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-bold text-indigo-700 cursor-pointer"
                  disabled={!mathFileAId}
                >
                  <option value="">-- Chọn khóa chính --</option>
                  {allAvailableFiles.find(f => f.id === mathFileAId)?.columns.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-slate-500 block uppercase">Khóa phụ (Tùy chọn):</label>
                <select
                  value={mathKeyA2}
                  onChange={(e) => setMathKeyA2(e.target.value)}
                  className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                  disabled={!mathFileAId}
                >
                  <option value="">-- Chọn khóa phụ --</option>
                  {allAvailableFiles.find(f => f.id === mathFileAId)?.columns.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* Tệp nguồn B */}
          <div className="space-y-3 p-4 bg-slate-50 rounded-xl border border-slate-200">
            <span className="text-[11px] font-bold text-amber-800 block uppercase tracking-wider flex items-center gap-1.5 border-b border-slate-200 pb-1.5">
              <span className="w-2.5 h-2.5 rounded-full bg-amber-500 block"></span> TỆP TIN NGUỒN B
            </span>
            
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div className="space-y-1">
                <label className="text-[10px] font-bold text-slate-500 block uppercase">Chọn tệp tin B:</label>
                <select
                  value={mathFileBId}
                  onChange={(e) => {
                    setMathFileBId(e.target.value);
                    setMathColB("");
                    setMathColB2("");
                    setMathColB3("");
                    setMathKeyB("");
                    setMathKeyB2("");
                  }}
                  className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                >
                  <option value="">-- Chọn tệp B --</option>
                  {allAvailableFiles.map(f => (
                    <option key={f.id} value={f.id}>{f.name.replace(/📂\s*\[TỆP CHÍNH\]\s*/, "")}</option>
                  ))}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-slate-500 block uppercase">Cột khóa chính (B):</label>
                <select
                  value={mathKeyB}
                  onChange={(e) => setMathKeyB(e.target.value)}
                  className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-bold text-amber-800 cursor-pointer"
                  disabled={!mathFileBId}
                >
                  <option value="">-- Chọn khóa chính --</option>
                  {allAvailableFiles.find(f => f.id === mathFileBId)?.columns.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-bold text-slate-500 block uppercase">Khóa phụ (Tùy chọn):</label>
                <select
                  value={mathKeyB2}
                  onChange={(e) => setMathKeyB2(e.target.value)}
                  className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                  disabled={!mathFileBId}
                >
                  <option value="">-- Chọn khóa phụ --</option>
                  {allAvailableFiles.find(f => f.id === mathFileBId)?.columns.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        </div>
      </fieldset>

      {/* 2. THIẾT LẬP PHÉP TÍNH CHÉO ĐA CỘT */}
      <fieldset className="border border-slate-200 rounded-xl p-4 sm:p-5 bg-white text-slate-800 space-y-4 shadow-sm relative">
        <legend className="px-3 text-xs sm:text-[12.5px] font-bold text-indigo-900 uppercase tracking-wide bg-slate-100 py-0.5 rounded border border-slate-200">
          2. THIẾT LẬP PHÉP TÍNH CHO CÁC CỘT CHỈ TIÊU SONG SONG
        </legend>

        <div className="space-y-3">
          {/* CẶP 1 */}
          <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-center bg-slate-50 p-3 sm:p-4 rounded-xl border border-slate-200">
            <div className="md:col-span-2 text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1">
              <CornerDownRight className="w-3.5 h-3.5 text-indigo-600 shrink-0" /> Cặp chỉ tiêu 1 <span className="text-rose-500">*</span>
            </div>
            <div className="md:col-span-3">
              <select
                value={mathColA}
                onChange={(e) => setMathColA(e.target.value)}
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                disabled={!mathFileAId}
              >
                <option value="">-- Chỉ tiêu tệp A --</option>
                {allAvailableFiles.find(f => f.id === mathFileAId)?.columns.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
            <div className="md:col-span-1.5 text-center">
              <select
                value={mathOp}
                onChange={(e) => setMathOp(e.target.value as any)}
                className="bg-indigo-900 border border-indigo-950 rounded px-2 py-1.5 text-xs font-black text-amber-400 text-center w-full shadow-sm cursor-pointer"
              >
                <option value="+">➕ Cộng</option>
                <option value="-">➖ Trừ</option>
                <option value="*">✖️ Nhân</option>
                <option value="/">➗ Chia</option>
              </select>
            </div>
            <div className="md:col-span-3">
              <select
                value={mathColB}
                onChange={(e) => setMathColB(e.target.value)}
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                disabled={!mathFileBId}
              >
                <option value="">-- Chỉ tiêu tệp B --</option>
                {allAvailableFiles.find(f => f.id === mathFileBId)?.columns.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
            <div className="md:col-span-2.5">
              <input
                type="text"
                value={mathNewColName}
                onChange={(e) => setMathNewColName(e.target.value)}
                placeholder="Tên cột kết quả 1..."
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-bold"
              />
            </div>
          </div>

          {/* CẶP 2 */}
          <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-center bg-slate-50/50 p-3 sm:p-4 rounded-xl border border-slate-200 border-dashed">
            <div className="md:col-span-2 text-xs font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
              <CornerDownRight className="w-3.5 h-3.5 text-slate-400 shrink-0" /> Cặp chỉ tiêu 2
            </div>
            <div className="md:col-span-3">
              <select
                value={mathColA2}
                onChange={(e) => setMathColA2(e.target.value)}
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                disabled={!mathFileAId}
              >
                <option value="">-- Chỉ tiêu tệp A (Tùy chọn) --</option>
                {allAvailableFiles.find(f => f.id === mathFileAId)?.columns.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
            <div className="md:col-span-1.5 text-center">
              <select
                value={mathOp2}
                onChange={(e) => setMathOp2(e.target.value as any)}
                className="bg-indigo-900/80 border border-indigo-950 rounded px-2 py-1.5 text-xs font-bold text-amber-300 text-center w-full cursor-pointer"
              >
                <option value="+">➕ Cộng</option>
                <option value="-">➖ Trừ</option>
                <option value="*">✖️ Nhân</option>
                <option value="/">➗ Chia</option>
              </select>
            </div>
            <div className="md:col-span-3">
              <select
                value={mathColB2}
                onChange={(e) => setMathColB2(e.target.value)}
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                disabled={!mathFileBId}
              >
                <option value="">-- Chỉ tiêu tệp B (Tùy chọn) --</option>
                {allAvailableFiles.find(f => f.id === mathFileBId)?.columns.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
            <div className="md:col-span-2.5">
              <input
                type="text"
                value={mathNewColName2}
                onChange={(e) => setMathNewColName2(e.target.value)}
                placeholder="Tên cột kết quả 2..."
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-bold"
              />
            </div>
          </div>

          {/* CẶP 3 */}
          <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-center bg-slate-50/50 p-3 sm:p-4 rounded-xl border border-slate-200 border-dashed">
            <div className="md:col-span-2 text-xs font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
              <CornerDownRight className="w-3.5 h-3.5 text-slate-400 shrink-0" /> Cặp chỉ tiêu 3
            </div>
            <div className="md:col-span-3">
              <select
                value={mathColA3}
                onChange={(e) => setMathColA3(e.target.value)}
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                disabled={!mathFileAId}
              >
                <option value="">-- Chỉ tiêu tệp A (Tùy chọn) --</option>
                {allAvailableFiles.find(f => f.id === mathFileAId)?.columns.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
            <div className="md:col-span-1.5 text-center">
              <select
                value={mathOp3}
                onChange={(e) => setMathOp3(e.target.value as any)}
                className="bg-indigo-900/80 border border-indigo-950 rounded px-2 py-1.5 text-xs font-bold text-amber-300 text-center w-full cursor-pointer"
              >
                <option value="+">➕ Cộng</option>
                <option value="-">➖ Trừ</option>
                <option value="*">✖️ Nhân</option>
                <option value="/">➗ Chia</option>
              </select>
            </div>
            <div className="md:col-span-3">
              <select
                value={mathColB3}
                onChange={(e) => setMathColB3(e.target.value)}
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-medium cursor-pointer"
                disabled={!mathFileBId}
              >
                <option value="">-- Chỉ tiêu tệp B (Tùy chọn) --</option>
                {allAvailableFiles.find(f => f.id === mathFileBId)?.columns.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
            <div className="md:col-span-2.5">
              <input
                type="text"
                value={mathNewColName3}
                onChange={(e) => setMathNewColName3(e.target.value)}
                placeholder="Tên cột kết quả 3..."
                className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 text-xs text-slate-800 focus:ring-1 focus:ring-indigo-500 font-bold"
              />
            </div>
          </div>
        </div>

        {/* Nút hành động chính */}
        <div className="pt-4 flex flex-col sm:flex-row items-center justify-between gap-4 border-t border-slate-100 mt-4">
          <div className="text-xs text-amber-800 bg-amber-50 px-3 py-2 rounded-lg border border-amber-200 font-medium font-sans">
            💡 <strong>Lưu ý:</strong> Sau khi đã chọn đầy đủ các tệp &amp; cột khóa ở Bước 1, vui lòng bấm nút bên phải để chạy phép toán so sánh!
          </div>
          <button
            onClick={handleExecuteWithLogs}
            className="w-full sm:w-auto bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-black text-xs sm:text-sm px-8 py-3.5 rounded-xl transition-all shadow-lg flex items-center justify-center gap-2 cursor-pointer uppercase tracking-widest active:scale-95 border-b-4 border-emerald-800"
          >
            <Play className="w-4 h-4 text-white shrink-0 animate-bounce" /> ⚡ BẤM ĐỂ CHẠY PHÉP TÍNH LIÊN KẾT CHÉO
          </button>
        </div>
      </fieldset>

      {/* 3. BẢNG SO SÁNH SONG SONG */}
      {(mathFileAId || mathFileBId) && (
        <fieldset className="border border-slate-200 rounded-xl p-4 sm:p-5 bg-white text-slate-800 space-y-4 shadow-sm relative">
          <legend className="px-3 text-xs sm:text-[12.5px] font-bold text-indigo-900 uppercase tracking-wide bg-slate-100 py-0.5 rounded border border-slate-200">
            3. THÔNG TIN ĐỐI CHIẾU SONG SONG PHÂN CẤP
          </legend>

          <p className="text-[11px] text-slate-500 leading-relaxed font-medium">
            Gõ bộ lọc từ khóa để tìm kiếm và gán nhanh cột liên kết dòng.
          </p>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            {/* Tệp A */}
            <div className="space-y-2">
              {mathFileAId ? (() => {
                const fileA = allAvailableFiles.find(f => f.id === mathFileAId);
                if (!fileA) return null;
                const filteredA = fileA.data.filter(row => {
                  if (!mathFilterA.trim()) return true;
                  return Object.values(row).some(v => 
                    String(v ?? "").toLowerCase().includes(mathFilterA.toLowerCase())
                  );
                });
                return (
                  <div className="bg-slate-50 p-3 rounded-lg border border-slate-200 space-y-2">
                    <div className="flex items-center justify-between gap-2 border-b border-slate-200 pb-1.5">
                      <span className="text-[11px] font-bold text-indigo-900 truncate max-w-[200px]" title={fileA.name}>
                        📂 Tệp A: {fileA.name.replace(/📂\s*\[TỆP CHÍNH\]\s*/, "")}
                      </span>
                      <span className="text-[10px] text-slate-500 font-mono font-bold bg-white px-2 py-0.5 rounded border border-slate-200">
                        {filteredA.length}/{fileA.data.length} dòng
                      </span>
                    </div>
                    <div className="flex gap-2">
                      <input
                        type="text"
                        placeholder="🔍 Tìm nhanh ngành tệp A (vd: Bán lẻ...)"
                        value={mathFilterA}
                        onChange={(e) => setMathFilterA(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded px-2.5 py-1 text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-medium"
                      />
                      {mathFilterA && (
                        <button 
                          onClick={() => setMathFilterA("")} 
                          className="text-[10px] text-slate-500 hover:text-slate-850 px-2 bg-slate-200 rounded font-bold"
                        >
                          Xóa
                        </button>
                      )}
                    </div>

                    <div className="overflow-x-auto max-h-[220px] overflow-y-auto border border-slate-200 rounded bg-white">
                      <table className="w-full text-left text-[10.5px]">
                        <thead className="bg-slate-100 text-slate-700 sticky top-0 font-bold border-b border-slate-200">
                          <tr>
                            {fileA.columns.slice(0, 5).map(col => {
                              const isKey = mathKeyA === col || mathKeyA2 === col;
                              const isVal = mathColA === col || mathColA2 === col || mathColA3 === col;
                              return (
                                <th 
                                  key={col} 
                                  className={`p-2 border-r border-slate-200 cursor-pointer hover:bg-slate-200 text-[10px] transition-all ${
                                    isKey ? "bg-indigo-50 text-indigo-800" : isVal ? "bg-amber-50 text-amber-800" : ""
                                  }`}
                                  onClick={() => {
                                    const act = confirm(`Bạn muốn gán cột "${col}" của Tệp A làm:\n- OK (Bấm OK): Làm cột số so sánh thứ 1\n- CANCEL (Bấm Hủy): Làm cột khóa liên kết chính`);
                                    if (act) {
                                      setMathColA(col);
                                      addLog(`Đã chọn nhanh cột [${col}] ở Tệp A làm Cột số tính toán 1.`, "info");
                                    } else {
                                      setMathKeyA(col);
                                      addLog(`Đã chọn nhanh cột [${col}] ở Tệp A làm Cột khóa liên kết dòng.`, "info");
                                    }
                                  }}
                                  title="Nhấp để gán cột"
                                >
                                  <span className="block truncate max-w-[100px]">{col}</span>
                                  <span className="block text-[8px] text-slate-500 font-bold">
                                    {isKey ? "🔑 Khóa" : isVal ? "🔢 Số" : "🖱️ Gán"}
                                  </span>
                                </th>
                              );
                            })}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 text-slate-800">
                          {filteredA.slice(0, 15).map((row, idx) => (
                            <tr key={idx} className="hover:bg-slate-50 border-b border-slate-100">
                              {fileA.columns.slice(0, 5).map(col => (
                                <td key={col} className="p-2 border-r border-slate-200 truncate max-w-[120px] font-medium" title={String(row[col] ?? "")}>
                                  {String(row[col] ?? "")}
                                </td>
                              ))}
                            </tr>
                          ))}
                          {filteredA.length === 0 && (
                            <tr>
                              <td colSpan={5} className="p-4 text-center text-slate-400">Không tìm thấy dòng phù hợp</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                );
              })() : (
                <div className="bg-slate-100 p-4 text-center rounded-lg border border-dashed border-slate-250 text-xs text-slate-400 py-10 font-bold">
                  Chưa chọn Tệp tin A.
                </div>
              )}
            </div>

            {/* Tệp B */}
            <div className="space-y-2">
              {mathFileBId ? (() => {
                const fileB = allAvailableFiles.find(f => f.id === mathFileBId);
                if (!fileB) return null;
                const filteredB = fileB.data.filter(row => {
                  if (!mathFilterB.trim()) return true;
                  return Object.values(row).some(v => 
                    String(v ?? "").toLowerCase().includes(mathFilterB.toLowerCase())
                  );
                });
                return (
                  <div className="bg-slate-50 p-3 rounded-lg border border-slate-200 space-y-2">
                    <div className="flex items-center justify-between gap-2 border-b border-slate-200 pb-1.5">
                      <span className="text-[11px] font-bold text-amber-800 truncate max-w-[200px]" title={fileB.name}>
                        📂 Tệp B: {fileB.name.replace(/📂\s*\[TỆP CHÍNH\]\s*/, "")}
                      </span>
                      <span className="text-[10px] text-slate-500 font-mono font-bold bg-white px-2 py-0.5 rounded border border-gray-200">
                        {filteredB.length}/{fileB.data.length} dòng
                      </span>
                    </div>
                    <div className="flex gap-2">
                      <input
                        type="text"
                        placeholder="🔍 Tìm nhanh ngành tệp B (vd: Bán lẻ...)"
                        value={mathFilterB}
                        onChange={(e) => setMathFilterB(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded px-2.5 py-1 text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-medium"
                      />
                      {mathFilterB && (
                        <button 
                          onClick={() => setMathFilterB("")} 
                          className="text-[10px] text-slate-500 hover:text-slate-855 px-2 bg-slate-200 rounded font-bold"
                        >
                          Xóa
                        </button>
                      )}
                    </div>

                    <div className="overflow-x-auto max-h-[220px] overflow-y-auto border border-slate-200 rounded bg-white">
                      <table className="w-full text-left text-[10.5px]">
                        <thead className="bg-slate-100 text-slate-700 sticky top-0 font-bold border-b border-slate-200">
                          <tr>
                            {fileB.columns.slice(0, 5).map(col => {
                              const isKey = mathKeyB === col || mathKeyB2 === col;
                              const isVal = mathColB === col || mathColB2 === col || mathColB3 === col;
                              return (
                                <th 
                                  key={col} 
                                  className={`p-2 border-r border-slate-200 cursor-pointer hover:bg-slate-200 text-[10px] transition-all ${
                                    isKey ? "bg-amber-50 text-amber-850" : isVal ? "bg-orange-50 text-orange-800" : ""
                                  }`}
                                  onClick={() => {
                                    const act = confirm(`Bạn muốn gán cột "${col}" của Tệp B làm:\n- OK (Bấm OK): Làm cột số so sánh thứ 1\n- CANCEL (Bấm Hủy): Làm cột khóa liên kết chính`);
                                    if (act) {
                                      setMathColB(col);
                                      addLog(`Đã chọn nhanh cột [${col}] ở Tệp B làm Cột số tính toán 1.`, "info");
                                    } else {
                                      setMathKeyB(col);
                                      addLog(`Đã chọn nhanh cột [${col}] ở Tệp B làm Cột khóa liên kết dòng.`, "info");
                                    }
                                  }}
                                  title="Nhấp để gán cột"
                                >
                                  <span className="block truncate max-w-[100px]">{col}</span>
                                  <span className="block text-[8px] text-slate-500 font-bold">
                                    {isKey ? "🔑 Khóa" : isVal ? "🔢 Số" : "🖱️ Gán"}
                                  </span>
                                </th>
                              );
                            })}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 text-slate-800">
                          {filteredB.slice(0, 15).map((row, idx) => (
                            <tr key={idx} className="hover:bg-slate-50 border-b border-slate-100">
                              {fileB.columns.slice(0, 5).map(col => (
                                <td key={col} className="p-2 border-r border-slate-200 truncate max-w-[120px] font-medium" title={String(row[col] ?? "")}>
                                  {String(row[col] ?? "")}
                                </td>
                              ))}
                            </tr>
                          ))}
                          {filteredB.length === 0 && (
                            <tr>
                              <td colSpan={5} className="p-4 text-center text-slate-400">Không tìm thấy dòng phù hợp</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                );
              })() : (
                <div className="bg-slate-100 p-4 text-center rounded-lg border border-dashed border-slate-250 text-xs text-slate-400 py-10 font-bold">
                  Chưa chọn Tệp tin B.
                </div>
              )}
            </div>
          </div>

          {/* PREVIEW KẾT QUẢ LIVE NHÁP */}
          {mathFileAId && mathFileBId && mathKeyA && mathKeyB && (
            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-2 mt-4">
              <span className="text-xs font-bold text-indigo-900 flex items-center gap-1.5 uppercase tracking-wide border-b border-slate-200 pb-1.5">
                📊 NHÁP XEM TRƯỚC THEO DÒNG KHỚP CỘT KHÓA (RESULTS PREVIEW)
              </span>

              {(() => {
                const fileA = allAvailableFiles.find(f => f.id === mathFileAId);
                const fileB = allAvailableFiles.find(f => f.id === mathFileBId);
                if (!fileA || !fileB) return null;

                const mapB = new Map<string, any>();
                fileB.data.forEach(row => {
                  let keyVal = String(row[mathKeyB] ?? "").trim().toLowerCase();
                  if (mathKeyB2) {
                    keyVal += "||" + String(row[mathKeyB2] ?? "").trim().toLowerCase();
                  }
                  if (keyVal !== "") {
                    mapB.set(keyVal, row);
                  }
                });

                const previewItems: any[] = [];
                fileA.data.forEach(rowA => {
                  let keyAVal = String(rowA[mathKeyA] ?? "").trim().toLowerCase();
                  if (mathKeyA2) {
                    keyAVal += "||" + String(rowA[mathKeyA2] ?? "").trim().toLowerCase();
                  }
                  const rowB = mapB.get(keyAVal);

                  const textMatch = !mathFilterA.trim() && !mathFilterB.trim() ? true :
                    String(rowA[mathKeyA] ?? "").toLowerCase().includes(mathFilterA.toLowerCase()) ||
                    (rowB && String(rowB[mathKeyB] ?? "").toLowerCase().includes(mathFilterB.toLowerCase()));

                  if (rowB && textMatch) {
                    previewItems.push({
                      key: rowA[mathKeyA],
                      key2: mathKeyA2 ? rowA[mathKeyA2] : null,
                      rowA,
                      rowB
                    });
                  }
                });

                if (previewItems.length === 0) {
                  return (
                    <p className="text-center text-xs text-rose-600 font-bold py-4">
                      ⚠️ Không tìm thấy dòng nào khớp trùng theo cột khóa đã chọn. Hãy kiểm tra lại cột khóa hoặc đổi từ khóa tìm kiếm chung.
                    </p>
                  );
                }

                return (
                  <div className="overflow-x-auto max-h-[220px] overflow-y-auto border border-slate-200 rounded bg-white">
                    <table className="w-full text-left text-[11px]">
                      <thead className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200 sticky top-0">
                        <tr>
                          <th className="p-2 border-r border-slate-200 text-xs">Cột Khóa liên kết</th>
                          {mathColA && (
                            <th className="p-2 border-r border-slate-200 text-indigo-900 text-right">
                              Tệp A: {mathColA}
                            </th>
                          )}
                          {mathColB && (
                            <th className="p-2 border-r border-slate-200 text-amber-900 text-right">
                              Tệp B: {mathColB}
                            </th>
                          )}
                          {mathColA && mathColB && (
                            <th className="p-2 border-r border-slate-200 text-emerald-800 font-bold text-right">
                              Công thức kết toán 1 ({mathOp})
                            </th>
                          )}
                          {mathColA2 && mathColB2 && (
                            <th className="p-2 text-purple-800 font-bold text-right">
                              Công thức kết toán 2 ({mathOp2})
                            </th>
                          )}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 text-slate-800 font-medium">
                        {previewItems.slice(0, 10).map((item, idx) => {
                          const valA = parseRobustNumber(item.rowA[mathColA]);
                          const valB = parseRobustNumber(item.rowB[mathColB]);
                          let res1 = 0;
                          if (mathOp === "+") res1 = valA + valB;
                          else if (mathOp === "-") res1 = valA - valB;
                          else if (mathOp === "*") res1 = valA * valB;
                          else if (mathOp === "/") res1 = valB !== 0 ? valA / valB : 0;
                          res1 = Math.round(res1 * 100) / 100;

                          let res2 = 0;
                          if (mathColA2 && mathColB2) {
                            const valA2 = parseRobustNumber(item.rowA[mathColA2]);
                            const valB2 = parseRobustNumber(item.rowB[mathColB2]);
                            if (mathOp2 === "+") res2 = valA2 + valB2;
                            else if (mathOp2 === "-") res2 = valA2 - valB2;
                            else if (mathOp2 === "*") res2 = valA2 * valB2;
                            else if (mathOp2 === "/") res2 = valB2 !== 0 ? valA2 / valB2 : 0;
                            res2 = Math.round(res2 * 100) / 100;
                          }

                          return (
                            <tr key={idx} className="hover:bg-slate-50 transition-all border-b border-slate-100">
                              <td className="p-2 border-r border-slate-200 font-bold text-slate-900 flex flex-col">
                                <span>{String(item.key)}</span>
                                {item.key2 && <span className="text-[9px] text-slate-500 font-semibold">({item.key2})</span>}
                              </td>
                              {mathColA && (
                                <td className="p-2 border-r border-slate-200 text-right text-indigo-800 font-mono font-bold">
                                  {valA.toLocaleString()}
                                </td>
                              )}
                              {mathColB && (
                                <td className="p-2 border-r border-slate-200 text-right text-amber-800 font-mono font-bold">
                                  {valB.toLocaleString()}
                                </td>
                              )}
                              {mathColA && mathColB && (
                                <td className="p-2 border-r border-slate-200 text-right text-emerald-800 font-black font-mono">
                                  {res1.toLocaleString()}
                                </td>
                              )}
                              {mathColA2 && mathColB2 && (
                                <td className="p-2 text-right text-purple-800 font-black font-mono">
                                  {res2.toLocaleString()}
                                </td>
                              )}
                            </tr>
                          );
                        })}
                        {previewItems.length > 10 && (
                          <tr className="bg-slate-100">
                            <td colSpan={5} className="p-2 text-center text-[10px] text-slate-500 italic font-bold">
                              Còn {previewItems.length - 10} dòng trùng khớp khác bên dưới... Bấm nút thực hiện phép tính phía trên để xử lý toàn bộ.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                );
              })()}
            </div>
          )}
        </fieldset>
      )}

      {/* 4. NHẬT KÝ ĐỒNG BỘ */}
      <div className="space-y-2">
        <span className="text-xs font-bold text-slate-600 tracking-wider uppercase font-mono block">
          Nhật ký hoạt động hệ thống
        </span>
        
        <div className="bg-slate-100 border border-slate-300 rounded-xl p-4 font-mono text-xs text-slate-800 h-[220px] overflow-y-auto shadow-inner relative border-b-4 border-slate-400">
          <div className="absolute top-2 right-4 flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
            <span className="text-[10px] text-emerald-700 font-bold uppercase tracking-widest scale-90">ONLINE</span>
          </div>

          <div className="space-y-1.5">
            {logs.map((log, idx) => {
              let typeColor = "text-slate-700";
              let prefix = "⚙️";
              if (log.type === "success") {
                typeColor = "text-emerald-700 font-bold";
                prefix = "✓ [THÀNH CÔNG]";
              } else if (log.type === "error") {
                typeColor = "text-rose-700 font-bold";
                prefix = "✗ [LỖI]";
              } else if (log.type === "warning") {
                typeColor = "text-amber-700 font-bold";
                prefix = "⚠ [CẢNH BÁO]";
              } else if (log.type === "info") {
                typeColor = "text-indigo-800 font-semibold";
                prefix = "ℹ [THÔNG TIN]";
              } else {
                prefix = "💻 [HỆ THỐNG]";
              }

              return (
                <div key={idx} className="flex items-start gap-2 border-b border-slate-200/50 pb-1 hover:bg-slate-200/30 transition-all leading-relaxed">
                  <span className="text-slate-450 shrink-0 select-none">[{log.time}]</span>
                  <span className={`${typeColor} break-all`}>
                    <span className="mr-1 opacity-90">{prefix}</span> {log.message}
                  </span>
                </div>
              );
            })}
            <div ref={logEndRef} />
          </div>
        </div>
        
        <p className="text-[10px] text-slate-500 italic font-sans flex items-center gap-1.5">
          <HelpCircle className="w-3.5 h-3.5 shrink-0 text-slate-400" /> Tiến trình xử lý dữ liệu được tự động cập nhật liên tục tại bảng nhật ký hoạt động.
        </p>
      </div>

    </div>
  );
}
