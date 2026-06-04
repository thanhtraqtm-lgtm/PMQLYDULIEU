import React, { useState, useMemo, useEffect } from "react";
import { 
  CheckSquare, 
  AlertTriangle, 
  XCircle, 
  CheckCircle2, 
  Sliders, 
  HelpCircle, 
  ArrowRightLeft, 
  Download, 
  Search, 
  Filter, 
  Info, 
  RefreshCw,
  Sparkles,
  FileSpreadsheet
} from "lucide-react";
import * as XLSX from "xlsx";

// 1. Loại bỏ dấu tiếng Việt để so khớp chính xác không phụ thuộc kiểu gõ
function removeVietnameseTones(str: string): string {
  if (!str) return "";
  str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
  str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
  str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
  str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
  str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
  str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
  str = str.replace(/đ/g, "d");
  str = str.replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, "A");
  str = str.replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, "E");
  str = str.replace(/Ì|Í|Ị|Ỉ|Ĩ/g, "I");
  str = str.replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, "O");
  str = str.replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, "U");
  str = str.replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, "Y");
  str = str.replace(/Đ/g, "D");
  // Các ký tự dấu tổ hợp khác
  str = str.replace(/\u0300|\u0301|\u0309|\u0303|\u0323/g, "");
  str = str.replace(/\u02C6|\u0306|\u031B/g, "");
  return str;
}

// 2. Tách từ & lọc bớt các từ "nhiễu" không mang nhiều ý nghĩa phân loại ngành
function getCleanWords(text: string): string[] {
  if (!text) return [];
  const normalized = removeVietnameseTones(text).toLowerCase();
  const cleanStr = normalized.replace(/[^a-z0-9\s]/g, " ");
  const words = cleanStr.split(/\s+/).filter(w => w.length > 0);
  
  const stopWords = new Set([
    "va", "cac", "dich", "vu", "ban", "hoat", "dong", "san", "pham", "kinh", "doanh", "cua", 
    "phuc", "vu", "tai", "cho", "do", "an", "uong", "nha", "hang", "le", "buon", "ho", "kinh", 
    "doanh", "dong", "co", "so", "nhom", "nganh", "tieu", "muc", "theo", "cap", "thi", "tu", 
    "cho", "biet"
  ]);
  
  return words.filter(w => !stopWords.has(w));
}

// Từ điển liên kết thông minh - Quy nạp ngữ nghĩa liên ngành của ĐTV và VSIC
const SYNONYM_MAP: { [key: string]: string[] } = {
  "bun": ["bun", "pho", "com", "an", "uong", "nha hang", "quan an", "lau", "nuong", "thit", "vit", "ga", "quay", "luoc", "chay"],
  "pho": ["bun", "pho", "com", "an", "uong", "nha hang", "quan an", "lau", "nuong", "thit", "vit", "ga", "quay", "luoc", "chay"],
  "com": ["bun", "pho", "com", "an", "uong", "nha hang", "quan an", "lau", "nuong", "thit", "vit", "ga", "quay", "luoc", "chay"],
  "lau": ["bun", "pho", "com", "an", "uong", "nha hang", "quan an", "lau", "nuong", "thit", "vit", "ga", "quay", "luoc", "chay"],
  "nuong": ["bun", "pho", "com", "an", "uong", "nha hang", "quan an", "lau", "nuong", "thit", "vit", "ga", "quay", "luoc", "chay"],
  "thit": ["bun", "pho", "com", "an", "uong", "nha hang", "quan an", "lau", "nuong", "thit", "vit", "ga", "quay", "luoc", "chay"],
  "cafe": ["cafe", "ca phe", "tra sua", "nuoc mia", "giai khat", "sinh to", "nuoc hoa qua", "giai khat", "quan nuoc", "do uong", "bia", "ruou"],
  "caphe": ["cafe", "ca phe", "tra sua", "nuoc mia", "giai khat", "sinh to", "nuoc hoa qua", "giai khat", "quan nuoc", "do uong", "bia", "ruou"],
  "sua": ["sua chua", "bao tri", "tivi", "tu lanh", "dien tu", "gia dung", "nghe nhin", "dien thoai", "may giat"],
  "dien": ["sua chua", "bao tri", "tivi", "tu lanh", "dien tu", "gia dung", "nghe nhin", "dien thoai", "may giat", "ngu kim", "kim khi", "sat", "thep", "day dien", "bong den"],
  "sat": ["ngu kim", "kim khi", "sat", "thep", "oc vit", "kim loai", "khoa cua", "phu tung"],
  "thue": ["thue", "cho thue", "tro", "nha", "can ho", "van hanh nha", "dat", "kiot", "mat bang", "kinh doanh"],
  "anh": ["anh", "chup", "photo", "studio", "nhiep anh", "quay phim", "anh ho so", "anh cuoi"],
  "chup": ["anh", "chup", "photo", "studio", "nhiep anh", "quay phim", "anh ho so", "anh cuoi"],
  "hoa": ["hoa tuoi", "cay canh", "hoa gia", "hat giong", "ca canh", "chim canh", "vat nuoi"],
  "nay": ["nail", "mong", "lam mong", "xam", "xam hinh", "tattoo", "spa", "massage", "tham my", "lam dep"],
  "mong": ["nail", "mong", "lam mong", "xam", "xam hinh", "tattoo", "spa", "massage", "tham my", "lam dep"],
  "xam": ["nail", "mong", "lam mong", "xam", "xam hinh", "tattoo", "spa", "massage", "tham my", "lam dep"],
  "tattoo": ["nail", "mong", "lam mong", "xam", "xam hinh", "tattoo", "spa", "massage", "tham my", "lam dep"]
};

// Hàm chấm điểm rà soát mô tả tương thích ngữ nghĩa
function evaluateSimilarity(descDTV: string, descStandard: string): { 
  score: number; 
  status: "SAFE" | "SUSPICIOUS" | "CRITICAL"; 
  reason: string;
} {
  const cleanDTV = (descDTV || "").trim();
  const cleanStandard = (descStandard || "").trim();

  if (!cleanDTV && !cleanStandard) {
    return { score: 100, status: "SAFE", reason: "Cả hai mô tả đều rỗng, bỏ qua rà soát." };
  }
  if (!cleanDTV || !cleanStandard) {
    return { score: 0, status: "CRITICAL", reason: "Một trong hai cột chứa dữ liệu trống hoàn toàn." };
  }

  const dtvNorm = removeVietnameseTones(cleanDTV).toLowerCase();
  const stdNorm = removeVietnameseTones(cleanStandard).toLowerCase();

  // 1. So sánh chính xác tương đồng tuyệt đối
  if (dtvNorm === stdNorm) {
    return { score: 100, status: "SAFE", reason: "Mô tả do ĐTV nhập trùng khớp 100% với tên danh mục chuẩn." };
  }

  // 2. Trích xuất các tập từ khóa có nghĩa
  const wordsDtv = getCleanWords(cleanDTV);
  const wordsStd = getCleanWords(cleanStandard);

  if (wordsDtv.length === 0 || wordsStd.length === 0) {
    return { score: 50, status: "SUSPICIOUS", reason: "Chuỗi ký tự chứa nhiều ký tự đặc biệt, chưa xác định từ khóa cốt lõi." };
  }

  // 3. Phân tích bẫy loại trừ đặc thù (Bẫy bốc mã lệch hẳn)
  const isDtvNailsOrTattoo = dtvNorm.includes("mong") || dtvNorm.includes("nail") || dtvNorm.includes("xam") || dtvNorm.includes("tattoo");
  const isStdHairOnly = stdNorm.includes("toc") && !stdNorm.includes("mong") && !stdNorm.includes("xam") && !stdNorm.includes("spa");
  
  if (isDtvNailsOrTattoo && isStdHairOnly) {
    return { 
      score: 15, 
      status: "SUSPICIOUS", 
      reason: `ĐTV khai báo cụ thể việc '${cleanDTV}' (làm móng/xăm hình) nhưng mã áp dụng lại thuộc nhóm '${cleanStandard}' (chỉ ghi nhận làm tóc).` 
    };
  }

  // 4. Đo mức độ giao thỏa từ khóa kết hợp từ điển đồng nghĩa (Synonyms)
  let matches = 0;
  const matchedWords: string[] = [];

  wordsDtv.forEach(wDtv => {
    // Nếu từ khóa có mặt nguyên mẫu trong ngành chuẩn
    if (wordsStd.includes(wDtv)) {
      matches++;
      matchedWords.push(wDtv);
      return;
    }

    // Nếu từ khóa có trong hệ thống ánh xạ từ đồng nghĩa
    const synonyms = Object.prototype.hasOwnProperty.call(SYNONYM_MAP, wDtv) ? SYNONYM_MAP[wDtv] : undefined;
    if (synonyms && Array.isArray(synonyms)) {
      const hasSynonymMatch = synonyms.some(syn => wordsStd.includes(syn) || stdNorm.includes(syn));
      if (hasSynonymMatch) {
         matches += 0.85; // Điểm quy nạp gián tiếp qua từ đồng nghĩa
         matchedWords.push(wDtv);
      }
    }
  });

  // Tính tỷ lệ phủ trùng khớp
  const dtvRatio = matches / wordsDtv.length;
  const stdRatio = matches / wordsStd.length;
  
  // Lấy giá trị tương đồng lớn nhất để bao quát (cho phép mô tả ngắn khớp chuẩn dài và ngược lại)
  const similarityScore = Math.round(Math.max(dtvRatio, stdRatio) * 100);

  // 5. Kết luận phân cấp dựa trên ngưỡng điểm
  if (similarityScore >= 45) {
    return { 
      score: similarityScore, 
      status: "SAFE", 
      reason: `Khớp an toàn (${similarityScore}%). Từ khóa liên thông tốt: [${matchedWords.join(", ")}]. Ngữ nghĩa tương thích cao.`
    };
  } else if (similarityScore >= 18) {
    return { 
      score: similarityScore, 
      status: "SUSPICIOUS", 
      reason: `Nghi ngờ sai sót (${similarityScore}%). Khớp quá ít từ khóa hoặc có sự sai lệch nhẹ về mặt phân ngành bổ trợ.` 
    };
  } else {
    return { 
      score: similarityScore, 
      status: "CRITICAL", 
      reason: `Sai lệch nghiêm trọng (${similarityScore}%). Mô tả thực tế '${cleanDTV}' không khớp bất kỳ mục từ hay nhóm đồng nghĩa nào của nhóm '${cleanStandard}'.` 
    };
  }
}

interface DescriptorMatchScannerProps {
  mainData: any[];
  columns: string[];
}

export default function DescriptorMatchScanner({ mainData, columns }: DescriptorMatchScannerProps) {
  // Trạng thái chọn cột
  const [colDTV, setColDTV] = useState<string>("");
  const [colStandard, setColStandard] = useState<string>("");
  const [strictness, setStrictness] = useState<number>(30); // Ngưỡng điểm nghi ngờ mặc định

  // Bộ lọc danh sách kết quả đối chiếu
  const [searchTerm, setSearchTerm] = useState<string>("");
  const [statusFilter, setStatusFilter] = useState<string>("ALL");

  // Kết quả sau khi chạy đối chiếu toàn cục
  const [scanResults, setScanResults] = useState<any[]>([]);
  const [isScanning, setIsScanning] = useState<boolean>(false);
  const [hasScanned, setHasScanned] = useState<boolean>(false);

  // Tự động nhận diện cấu trúc cột tối ưu lúc khởi chạy
  useEffect(() => {
    if (columns && columns.length > 0) {
      // Tìm cột ĐTV mô tả tự động
      const autoDTV = columns.find(c => {
         const name = c.toLowerCase();
         return name.includes("mô tả") || name.includes("mo ta") || name.includes("hoạt động") || name.includes("hoat dong") || name.includes("điều tra viên") || name.includes("dtv");
      });
      // Tìm cột ngành chuẩn
      const autoStd = columns.find(c => {
         const name = c.toLowerCase();
         return name.includes("chuẩn") || name.includes("chuan") || name.includes("tên ngành") || name.includes("ten nganh") || name.includes("phân cấp") || name.includes("cấp 5") || name.includes("lộ trình");
      });

      if (autoDTV) setColDTV(autoDTV);
      else setColDTV(columns[0] || "");

      if (autoStd) setColStandard(autoStd);
      else if (columns.length > 1) setColStandard(columns[1]);
      else if (columns.length > 0) setColStandard(columns[0]);
    }
  }, [columns]);

  // Tiến hành chạy thuật toán rà soát đối chiếu thông minh cho 10,000 dòng
  const handleRunMatchScan = () => {
    if (!colDTV || !colStandard) {
      alert("Vui lòng chỉ định rõ Cột mô tả của ĐTV và Cột tên ngành chuẩn để phần mềm phân tích!");
      return;
    }

    setIsScanning(true);
    setHasScanned(false);

    // Sử dụng setTimeout để không treo luồng UI đối với số lượng bản ghi cực lớn (10,000+ dòng)
    setTimeout(() => {
      try {
        const analyzed = mainData.map((row, index) => {
          if (!row || typeof row !== 'object') {
            return {
              index: index + 1,
              originalRow: {},
              descDTV: "",
              descStandard: "",
              score: 0,
              status: "SAFE",
              reason: "Dòng trống"
            };
          }
          const valDtv = String(row[colDTV] || "").trim();
          const valStd = String(row[colStandard] || "").trim();
          
          const evaluation = evaluateSimilarity(valDtv, valStd);
          
          // Ghi đè trạng thái nếu điểm số nằm dưới ngưỡng thiết lập của người dùng
          let finalStatus = evaluation.status;
          if (finalStatus === "SAFE" && evaluation.score < strictness) {
            finalStatus = "SUSPICIOUS";
          }

          return {
            index: index + 1,
            originalRow: row,
            descDTV: valDtv,
            descStandard: valStd,
            score: evaluation.score,
            status: finalStatus,
            reason: evaluation.reason
          };
        });

        setScanResults(analyzed);
        setHasScanned(true);
      } catch (err: any) {
        alert("Lỗi rà soát hệ thống: " + err.message);
      } finally {
        setIsScanning(false);
      }
    }, 150);
  };

  // Thống kê phân loại dữ liệu sau rà quét
  const stats = useMemo(() => {
    if (!hasScanned) return { total: 0, safe: 0, suspicious: 0, critical: 0 };
    
    const total = scanResults.length;
    const safe = scanResults.filter(r => r.status === "SAFE").length;
    const suspicious = scanResults.filter(r => r.status === "SUSPICIOUS").length;
    const critical = scanResults.filter(r => r.status === "CRITICAL").length;

    return { total, safe, suspicious, critical };
  }, [scanResults, hasScanned]);

  // Lọc dữ liệu hiển thị theo điều kiện nhập và tab lọc
  const filteredScanItems = useMemo(() => {
    if (!hasScanned) return [];
    
    return scanResults.filter(item => {
      // 1. Lọc theo trạng thái badge
      if (statusFilter !== "ALL" && item.status !== statusFilter) {
        return false;
      }
      // 2. Lọc theo chuỗi tìm kiếm
      if (searchTerm.trim()) {
        const query = searchTerm.toLowerCase();
        return (
          item.descDTV.toLowerCase().includes(query) ||
          item.descStandard.toLowerCase().includes(query) ||
          item.reason.toLowerCase().includes(query) ||
          String(item.index).includes(query)
        );
      }
      return true;
    });
  }, [scanResults, hasScanned, statusFilter, searchTerm]);

  // Giới hạn để render tối ưu mượt mà cho bảng
  const displayLimit = 200;
  const displayedScanItems = filteredScanItems.slice(0, displayLimit);

  // Kết xuất Excel báo cáo đối chiếu thông minh
  const handleExportMatchReport = () => {
    if (scanResults.length === 0) {
      alert("Chưa có kết quả phân tích phục vụ kết xuất!");
      return;
    }

    try {
      const wb = XLSX.utils.book_new();

      // Cấu trúc lại các dòng để xuất một bảng biểu khoa học
      const sheetData = scanResults.map(item => {
        let textStatus = "AN TOÀN";
        if (item.status === "SUSPICIOUS") textStatus = "NGHI NGỜ SAI LỆCH";
        if (item.status === "CRITICAL") textStatus = "SAI SAI BIỆT NẶNG";

        // Trích xuất toàn bộ thuộc tính gốc của hàng, sau đó dán thêm 3 cột đánh giá cuối cùng
        return {
          ...item.originalRow,
          "Mô tả ĐTV Chỉ định": item.descDTV,
          "Tên ngành chuẩn VSIC": item.descStandard,
          "Điểm Tương Đồng (%)": item.score,
          "Trạng Thái Khớp Mô Tả": textStatus,
          "Lý Giải / Đề Xuất Khắc Phục": item.reason
        };
      });

      const ws = XLSX.utils.json_to_sheet(sheetData);
      XLSX.utils.book_append_sheet(wb, ws, "Báo_Cáo_Đối_Chiếu_Mô_Tả");
      XLSX.writeFile(wb, "Bao_Cao_Doi_Chieu_Mo_Ta_Hoat_Dong_VSIC.xlsx");
      
      alert("Đã kết xuất báo cáo Excel thành công! Tệp tin chứa đầy đủ các thuộc tính gốc cùng với các cột Phân Loại Tương Đồng, Điểm Số và Diễn Giải chi tiết tương thích.");
    } catch (e: any) {
      alert("Lỗi kết tinh Excel: " + e.message);
    }
  };

  return (
    <div className="bg-[#1f2937] border border-[#374151] rounded-2xl p-6 space-y-6">
      
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-gray-800 pb-5">
        <div className="space-y-1">
          <h3 className="text-lg font-bold text-white flex items-center gap-2">
            <ArrowRightLeft className="w-5 h-5 text-purple-400" /> CÔNG CỤ ĐỐI CHIẾU MÔ TẢ ĐTV VS. TÊN NGÀNH CHUẨN (HỌC SEMANTIC)
          </h3>
          <p className="text-xs text-gray-400">
            Hệ thống rà quét tự động song song và so sánh mức tương thích ngữ nghĩa từ vựng giữa Cột mô tả thực tế của Điều tra viên nhập và Cột tên ngành cấp 5 chuẩn để phát hiện sai lệch ngầm hoặc sai mã.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <div className="bg-purple-600/10 border border-purple-500/20 text-purple-400 px-3 py-1.5 rounded-xl font-mono text-xs flex items-center gap-1.5">
            <Sparkles className="w-3.5 h-3.5" /> Chạy Offline cục bộ: Siêu nhanh
          </div>
        </div>
      </div>

      {/* 1. Thiết lập chọn cột rà soát */}
      <div className="bg-[#111827] rounded-xl p-5 border border-purple-950/20 space-y-4">
        <h4 className="text-xs font-bold text-gray-400 uppercase tracking-widest flex items-center gap-1.5 pb-2 border-b border-gray-800">
          <Sliders className="w-4 h-4 text-purple-400" /> CÀI ĐẶT BÀN CHỈ ĐỊNH CỘT ĐỐI CHIẾU
        </h4>
        
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          
          {/* Cột A: Mô tả ĐTV */}
          <div className="space-y-1.5">
            <label className="text-xs text-gray-300 font-semibold flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-amber-400"></span>
              Cột Mô tả của ĐTV nhập (Cột so sánh gốc):
            </label>
            <select
              value={colDTV}
              onChange={(e) => setColDTV(e.target.value)}
              className="w-full bg-[#1f2937] border border-[#374151] rounded-xl px-3 py-2.5 text-xs text-white placeholder-gray-400 focus:ring-1 focus:ring-purple-500 outline-none"
            >
              <option value="">-- Chọn cột mô tả --</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
            <p className="text-[10px] text-gray-500">Ví dụ: Mô tả hoạt động kinh doanh (quán bún chả, bán nước...)</p>
          </div>

          {/* Cột B: Tên chuẩn ngành */}
          <div className="space-y-1.5">
            <label className="text-xs text-gray-300 font-semibold flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-emerald-400"></span>
              Cột Tên ngành VSIC chuẩn (Dữ liệu đối sánh):
            </label>
            <select
              value={colStandard}
              onChange={(e) => setColStandard(e.target.value)}
              className="w-full bg-[#1f2937] border border-[#374151] rounded-xl px-3 py-2.5 text-xs text-white placeholder-gray-400 focus:ring-1 focus:ring-purple-500 outline-none"
            >
              <option value="">-- Chọn cột tên ngành chuẩn --</option>
              {columns.map(col => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
            <p className="text-[10px] text-gray-500">Ví dụ: Tên ngành cấp 5 chuẩn từ Catalog VSIC tương ứng mã.</p>
          </div>

          {/* Độ nhạy rà soát */}
          <div className="space-y-1.5">
            <label className="text-xs text-gray-300 font-semibold justify-between flex items-center">
              <span>Độ nhạy quét (Bẫy từ khóa lệch):</span>
              <span className="text-purple-400 font-mono font-bold text-xs">{strictness}% tương đồng</span>
            </label>
            <div className="pt-2">
              <input 
                type="range" 
                min="10" 
                max="70" 
                step="5"
                value={strictness}
                onChange={(e) => setStrictness(Number(e.target.value))}
                className="w-full h-1.5 bg-gray-700 rounded-lg appearance-none cursor-pointer accent-purple-500"
              />
              <div className="flex justify-between text-[10px] text-gray-500 pt-1 font-mono">
                <span>Rộng rãi (10%)</span>
                <span>Tiêu chuẩn (30%)</span>
                <span>Khắt khe (70%)</span>
              </div>
            </div>
          </div>

        </div>

        {/* Action Buttons */}
        <div className="flex flex-col sm:flex-row gap-3 pt-2">
          <button
            onClick={handleRunMatchScan}
            disabled={isScanning}
            className="flex-1 bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white font-bold text-xs py-3 px-6 rounded-xl transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer disabled:opacity-50"
          >
            {isScanning ? (
              <>
                <RefreshCw className="w-4 h-4 animate-spin text-white" /> ĐANG PHÂN TÍCH QUY NẠP NGỮ NGHĨA {mainData.length} DÒNG...
              </>
            ) : (
              <>
                <CheckSquare className="w-4 h-4 text-purple-200" /> BẮT ĐẦU CHẠY SO SÁNH & ĐÁNH DẤU CHÊNH LỆCH
              </>
            )}
          </button>

          {hasScanned && (
            <button
              onClick={handleExportMatchReport}
              className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs py-3 px-6 rounded-xl transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer"
            >
              <Download className="w-4 h-4" /> XUẤT BÁO CÁO ĐỐI CHIẾU (.XLSX)
            </button>
          )}
        </div>
      </div>

      {/* 2. Hiển thị thông số dạng Bento Grid khi đã rà soát */}
      {hasScanned && (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          
          {/* Tổng kiểm rà */}
          <div className="bg-[#111827] border border-gray-800 rounded-2xl p-4 flex items-center justify-between">
            <div className="space-y-0.5">
              <span className="text-[10px] text-gray-500 font-bold uppercase tracking-wider">Tổng rà quét</span>
              <p className="text-xl font-black text-white font-mono">{stats.total.toLocaleString()}</p>
            </div>
            <div className="bg-blue-500/10 p-2.5 rounded-xl border border-blue-500/20">
              <Info className="w-5 h-5 text-blue-400" />
            </div>
          </div>

          {/* Khớp an toàn */}
          <div className="bg-[#111827] border border-gray-800 rounded-2xl p-4 flex items-center justify-between">
            <div className="space-y-0.5">
              <span className="text-[10px] text-emerald-500 font-bold uppercase tracking-wider">Khớp an toàn</span>
              <p className="text-xl font-black text-emerald-400 font-mono">
                {stats.safe.toLocaleString()} 
                <span className="text-xs text-gray-500 ml-1 font-normal">({stats.total > 0 ? Math.round(stats.safe / stats.total * 100) : 0}%)</span>
              </p>
            </div>
            <div className="bg-emerald-500/10 p-2.5 rounded-xl border border-emerald-500/20">
              <CheckCircle2 className="w-5 h-5 text-emerald-400" />
            </div>
          </div>

          {/* Nghi ngờ sai lệch */}
          <div className="bg-[#111827] border border-gray-800 rounded-2xl p-4 flex items-center justify-between">
            <div className="space-y-0.5">
              <span className="text-[10px] text-amber-500 font-bold uppercase tracking-wider">Nghi ngờ sai lệch</span>
              <p className="text-xl font-black text-amber-400 font-mono">
                {stats.suspicious.toLocaleString()} 
                <span className="text-xs text-gray-500 ml-1 font-normal">({stats.total > 0 ? Math.round(stats.suspicious / stats.total * 100) : 0}%)</span>
              </p>
            </div>
            <div className="bg-amber-500/10 p-2.5 rounded-xl border border-amber-500/20">
              <AlertTriangle className="w-5 h-5 text-amber-400" />
            </div>
          </div>

          {/* Sai số nghiêm trọng */}
          <div className="bg-[#111827] border border-gray-800 rounded-2xl p-4 flex items-center justify-between">
            <div className="space-y-0.5">
              <span className="text-[10px] text-rose-500 font-bold uppercase tracking-wider">Sai biệt nặng</span>
              <p className="text-xl font-black text-rose-400 font-mono">
                {stats.critical.toLocaleString()} 
                <span className="text-xs text-gray-500 ml-1 font-normal">({stats.total > 0 ? Math.round(stats.critical / stats.total * 100) : 0}%)</span>
              </p>
            </div>
            <div className="bg-rose-500/10 p-2.5 rounded-xl border border-rose-500/20">
              <XCircle className="w-5 h-5 text-rose-400" />
            </div>
          </div>

        </div>
      )}

      {/* 3. Danh sách kết quả rà quét chi tiết kèm theo Tìm kiếm và Bộ lọc */}
      {hasScanned && (
        <div className="space-y-4">
          
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            
            {/* Bộ lọc Tab */}
            <div className="flex flex-wrap items-center gap-1.5 bg-[#111827] p-1.5 rounded-xl border border-gray-800">
              <button
                onClick={() => setStatusFilter("ALL")}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                  statusFilter === "ALL" ? "bg-purple-600 text-white" : "text-gray-400 hover:text-white"
                }`}
              >
                Tất cả ({stats.total})
              </button>
              <button
                onClick={() => setStatusFilter("SAFE")}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1 ${
                  statusFilter === "SAFE" ? "bg-emerald-600 text-white" : "text-gray-400 hover:text-emerald-400"
                }`}
              >
                💚 An toàn ({stats.safe})
              </button>
              <button
                onClick={() => setStatusFilter("SUSPICIOUS")}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1 ${
                  statusFilter === "SUSPICIOUS" ? "bg-amber-600 text-white" : "text-gray-400 hover:text-amber-400"
                }`}
              >
                💛 Nghi ngờ ({stats.suspicious})
              </button>
              <button
                onClick={() => setStatusFilter("CRITICAL")}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1 ${
                  statusFilter === "CRITICAL" ? "bg-rose-600 text-white" : "text-gray-400 hover:text-rose-400"
                }`}
              >
                💔 Sai biệt nặng ({stats.critical})
              </button>
            </div>

            {/* Tìm kiếm nhanh */}
            <div className="relative w-full md:w-[280px]">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-gray-500">
                <Search className="w-4 h-4" />
              </span>
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Tìm text mô tả hoặc từ khóa..."
                className="w-full bg-[#111827] border border-gray-800 focus:border-purple-500 focus:ring-1 focus:ring-purple-500 rounded-xl pl-9 pr-4 py-2 text-xs text-white placeholder-gray-500 outline-none"
              />
            </div>

          </div>

          {/* Bảng dữ liệu tương tác */}
          <div className="border border-gray-800 rounded-2xl overflow-hidden bg-[#111827]/60">
            <div className="max-h-[550px] overflow-y-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-[#111827] border-b border-gray-800 sticky top-0 z-10 text-gray-400 font-mono text-[10px] tracking-wider uppercase">
                    <th className="p-3.5 text-center w-[75px]">STT Hàng</th>
                    <th className="p-3.5 min-w-[200px]">MÔ TẢ CỦA ĐTV KHAI BÁO (CỘT A)</th>
                    <th className="p-3.5 min-w-[200px]">TÊN NGÀNH CHUẨN VSIC CẤP 5 (CỘT B)</th>
                    <th className="p-3.5 text-center w-[110px]">TỈ LỆ TRÙNG</th>
                    <th className="p-3.5">PHÂN TÍCH LOGIC & ĐỀ XUẤT ĐỐI CHIẾU BIỂU</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-850/40 text-xs">
                  {displayedScanItems.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="p-10 text-center text-gray-500 italic">
                        Không tìm thấy dòng đánh giá nào thích khớp với bộ lọc tìm kiếm hiện hành.
                      </td>
                    </tr>
                  ) : (
                    displayedScanItems.map((item, idx) => {
                      let statusBadge = null;
                      let trBg = "";

                      switch (item.status) {
                        case "SAFE":
                          statusBadge = (
                            <span className="inline-flex items-center gap-1 px-2.5 py-1 bg-emerald-950/40 text-emerald-400 border border-emerald-500/10 rounded-lg text-[10px] font-bold">
                              💚 AN TOÀN
                            </span>
                          );
                          break;
                        case "SUSPICIOUS":
                          statusBadge = (
                            <span className="inline-flex items-center gap-1 px-2.5 py-1 bg-amber-950/40 text-amber-400 border border-amber-500/10 rounded-lg text-[10px] font-bold">
                              💛 NGHI NGỜ
                            </span>
                          );
                          trBg = "bg-amber-950/5 hover:bg-amber-950/10";
                          break;
                        case "CRITICAL":
                          statusBadge = (
                            <span className="inline-flex items-center gap-1 px-2.5 py-1 bg-rose-950/40 text-rose-400 border border-rose-500/10 rounded-lg text-[10px] font-bold">
                              💔 SAI KHÁC NẶNG
                            </span>
                          );
                          trBg = "bg-rose-950/5 hover:bg-rose-950/10";
                          break;
                      }

                      return (
                        <tr key={item.index} className={`${trBg} transition-colors hover:bg-gray-800/20`}>
                          <td className="p-3.5 text-center font-mono font-bold text-gray-500">
                            {item.index}
                          </td>
                          <td className="p-3.5 font-medium text-amber-300 font-sans leading-relaxed">
                            {item.descDTV}
                          </td>
                          <td className="p-3.5 text-gray-100 font-sans leading-relaxed">
                            {item.descStandard}
                          </td>
                          <td className="p-3.5 text-center">
                            <span className={`font-mono font-extrabold text-sm ${
                              item.score >= 50 ? "text-emerald-400" : item.score >= 20 ? "text-amber-400" : "text-rose-400"
                            }`}>
                              {item.score}%
                            </span>
                          </td>
                          <td className="p-3.5 space-y-1.5 leading-normal">
                            <div className="flex items-center gap-2 flex-wrap">
                              {statusBadge}
                            </div>
                            <p className="text-gray-400 text-[11px] font-sans">
                              {item.reason}
                            </p>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
            
            {filteredScanItems.length > displayLimit && (
              <div className="p-4 bg-[#111827] border-t border-gray-800 text-center text-xs text-gray-500 italic">
                Hiển thị tối đa {displayLimit} kết quả lọc tiêu biểu nhất trên trình duyệt để tránh tràn bộ nhớ. Vui lòng nhấn nút "Xuất báo cáo đối chiếu" để tải về bản Excel chứa đầy đủ {filteredScanItems.length} dòng đánh giá rà soát!
              </div>
            )}
          </div>

        </div>
      )}

      {/* 4. Giới thiệu quy chuẩn mẫu hướng dẫn người dùng khi chưa chạy rà quét */}
      {!hasScanned && (
        <div className="bg-[#111827] rounded-xl p-6 border border-purple-500/15 space-y-4">
          <div className="flex items-center gap-2">
            <HelpCircle className="w-5 h-5 text-purple-400" />
            <h4 className="text-sm font-bold text-white uppercase tracking-wider">Cơ chế hoạt động & Hướng dẫn sử dụng mẫu:</h4>
          </div>
          
          <div className="space-y-4 text-xs text-gray-300 leading-relaxed">
            <p className="text-gray-400">
              Công cụ này so sánh sự trùng khớp ngữ nghĩa từ vựng tiếng Việt giữa cột <strong>Mô tả của Điều tra viên</strong> (A) và cột <strong>Tên gọi mặc định của Mã ngành</strong> (B) dựa trên giải thuật bóc tách từ chuẩn (Stemming) kết hợp từ điển ánh xạ nhóm liên thông:
            </p>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-[#192130]/60 p-4 rounded-xl border border-gray-800">
              <div>
                <span className="text-emerald-400 font-bold">✓ Trường hợp Khớp / An toàn:</span>
                <ul className="list-disc list-inside mt-2 space-y-1 text-gray-400 pl-1">
                  <li><strong className="text-white">"quán bún chả nướng"</strong> tương đồng cao với nhóm ngành chuẩn <em className="text-amber-400">"Hàng ăn uống, nhà hàng..."</em>.</li>
                  <li><strong className="text-white">"bán đồ điện hàng sắt"</strong> khớp an toàn với nhóm đồ ngũ kim bán lẻ.</li>
                  <li><strong className="text-white">"cho thuê nhà kinh doanh"</strong> liên kết trực tiếp với dịch vụ bất động sản.</li>
                </ul>
              </div>
              
              <div>
                <span className="text-amber-400 font-bold">⚠ Trường hợp cảnh báo Nghi ngờ / Sai lệch:</span>
                <ul className="list-disc list-inside mt-2 space-y-1 text-gray-400 pl-1">
                  <li><strong className="text-white">"dịch vụ làm móng, săm hình"</strong> gán vào ngành chuẩn <em className="text-yellow-400">"Dịch vụ làm tóc"</em> (Làm móng, xăm hình thuộc phân nhóm thẩm mỹ cơ thể khác, bốc mã làm tóc là nghi ngờ).</li>
                  <li>Cột mô tả ghi <strong className="text-white">"Dịch vụ giặt là"</strong> mà cột ngành chuẩn ghi <em className="text-rose-400">"Bán buôn vải sợi"</em> (Tỉ lệ trùng khớp 0% kịch khung).</li>
                </ul>
              </div>
            </div>

            <p className="font-mono text-[10px] text-purple-400/80 italic">
              * Mẹo: Bạn chỉ cần tải tệp dữ liệu lên ứng dụng, chọn cột phù hợp đại diện cho 2 trường này trên bảng thiết lập phía trên, điều chỉnh độ nhạy rà soát rồi nhấp nút chạy là xong!
            </p>
          </div>
        </div>
      )}

    </div>
  );
}
