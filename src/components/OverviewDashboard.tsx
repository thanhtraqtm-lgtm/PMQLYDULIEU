import React, { useState } from "react";
import {
  Sparkles,
  Zap,
  CheckCircle2,
  AlertTriangle,
  FileSpreadsheet,
  UploadCloud,
  ArrowRight,
  Play,
  RotateCcw,
  Search,
  Scale,
  BarChart3,
  Bot,
  Check
} from "lucide-react";

// Dữ liệu mẫu thực tế mô phỏng 30 dòng điều tra thống kê
export const SAMPLE_DATASET = [
  {
    "STT": 1,
    "Tên cơ sở": "Xưởng cơ khí Hoàng Phát",
    "Mô tả ngành nghề": "Gia công hàn tiện cơ khí, làm cửa sắt, mái tôn",
    "Mã ngành": "47110", // Lệch: 47110 là Bán lẻ, đúng phải là 25920
    "Số lao động": 5,
    "Doanh thu (triệu đồng)": 1250,
    "Mã xã phường": "001"
  },
  {
    "STT": 2,
    "Tên cơ sở": "Tạp hóa Minh Anh",
    "Mô tả ngành nghề": "Bán lẻ bánh kẹo, bia nước ngọt, nhu yếu phẩm gia đình",
    "Mã ngành": "47110", // Đúng
    "Số lao động": 2,
    "Doanh thu (triệu đồng)": 480,
    "Mã xã phường": "002"
  },
  {
    "STT": 3,
    "Tên cơ sở": "Hộ may mặc Thu Hà",
    "Mô tả ngành nghề": "Gia công may đo quần áo, sửa chữa trang phục",
    "Mã ngành": "14100", // Đúng
    "Số lao động": 4,
    "Doanh thu (triệu đồng)": 360,
    "Mã xã phường": "001"
  },
  {
    "STT": 4,
    "Tên cơ sở": "DNTN Vận tải & Thương mại Hải Đăng",
    "Mô tả ngành nghề": "Vận tải hàng hóa bằng xe tải liên tỉnh",
    "Mã ngành": "49339", // Đúng
    "Số lao động": 1,
    "Doanh thu (triệu đồng)": 85000, // Lệch logic cực đoan: 1 người mà 85 tỷ!
    "Mã xã phường": "003"
  },
  {
    "STT": 5,
    "Tên cơ sở": "Quán Cơm tấm Sài Gòn",
    "Mô tả ngành nghề": "Phục vụ ăn uống tại chỗ, cơm phần, bún phở",
    "Mã ngành": "56101", // Đúng
    "Số lao động": 6,
    "Doanh thu (triệu đồng)": 720,
    "Mã xã phường": "999" // Lệch: Mã xã không có trong danh mục
  },
  {
    "STT": 6,
    "Tên cơ sở": "Nhà thuốc An Khang",
    "Mô tả ngành nghề": "Bán lẻ thuốc tân dược, dụng cụ y tế gia đình",
    "Mã ngành": "47721",
    "Số lao động": 3,
    "Doanh thu (triệu đồng)": 1500,
    "Mã xã phường": "002"
  },
  {
    "STT": 7,
    "Tên cơ sở": "Garage Sửa chữa ô tô Thành Đạt",
    "Mô tả ngành nghề": "Bảo dưỡng, sửa chữa ô tô và xe có động cơ khác",
    "Mã ngành": "45110", // Lệch: 45110 là Bán buôn ô tô, đúng phải là 45200
    "Số lao động": 8,
    "Doanh thu (triệu đồng)": 2400,
    "Mã xã phường": "001"
  },
  {
    "STT": 8,
    "Tên cơ sở": "Tiệm làm tóc & Spa Thảo Vy",
    "Mô tả ngành nghề": "Cắt uốn tóc, gội đầu, chăm sóc da mặt",
    "Mã ngành": "96310",
    "Số lao động": 4,
    "Doanh thu (triệu đồng)": 310,
    "Mã xã phường": "003"
  },
  {
    "STT": 9,
    "Tên cơ sở": "Đại lý Thức ăn gia súc Bình Minh",
    "Mô tả ngành nghề": "Bán buôn cám gia súc gia cầm, thức ăn chăn nuôi",
    "Mã ngành": "46209",
    "Số lao động": 3,
    "Doanh thu (triệu đồng)": 5600,
    "Mã xã phường": "002"
  },
  {
    "STT": 10,
    "Tên cơ sở": "Cửa hàng Điện thoại Di động Tiến Lực",
    "Mô tả ngành nghề": "Bán lẻ điện thoại di động, máy tính bảng và phụ kiện",
    "Mã ngành": "47412",
    "Số lao động": 2,
    "Doanh thu (triệu đồng)": 1800,
    "Mã xã phường": "001"
  }
];

// 3 Kịch bản quét mẫu gọn gàng
interface ScanScenario {
  id: string;
  tag: string;
  tagColor: string;
  title: string;
  facility: string;
  desc: string;
  code: string;
  workers: number;
  revenue: string;
  errorField: "code" | "math" | "admin";
  aiVerdict: {
    status: "error" | "warning";
    title: string;
    detected: string;
    suggest: string;
    targetValue?: string;
  };
}

const SCAN_SCENARIOS: ScanScenario[] = [
  {
    id: "vsic-mismatch",
    tag: "LỆCH MÃ NGÀNH",
    tagColor: "bg-rose-50 text-rose-700 border-rose-200",
    title: "Cơ khí nhưng điền mã Bán lẻ",
    facility: "Xưởng cơ khí Hoàng Phát",
    desc: "Gia công hàn tiện cơ khí, làm cửa sắt, mái tôn",
    code: "47110",
    workers: 5,
    revenue: "1.250 triệu",
    errorField: "code",
    aiVerdict: {
      status: "error",
      title: "MÃ NGÀNH KHÔNG KHỚP MÔ TẢ",
      detected: "Mã 47110 là Bán lẻ siêu thị, mô tả thực tế là cơ khí.",
      suggest: "Đổi sang mã 25920 (Gia công cơ khí, xử lý kim loại).",
      targetValue: "25920"
    }
  },
  {
    id: "math-outlier",
    tag: "BẤT THƯỜNG TOÁN",
    tagColor: "bg-amber-50 text-amber-800 border-amber-200",
    title: "1 người khai 85 tỷ đồng",
    facility: "DNTN Vận tải Hải Đăng",
    desc: "Vận tải hàng hóa bằng xe tải liên tỉnh",
    code: "49339",
    workers: 1,
    revenue: "85.000 triệu",
    errorField: "math",
    aiVerdict: {
      status: "warning",
      title: "DOANH THU BẤT THƯỜNG",
      detected: "Tỷ suất 85 tỷ / người vượt 120 lần trung bình ngành.",
      suggest: "Khả năng nhầm đơn vị tính hoặc ghi nhầm cột lao động.",
      targetValue: "850 triệu"
    }
  },
  {
    id: "admin-code",
    tag: "SAI MÃ XÃ",
    tagColor: "bg-sky-50 text-sky-800 border-sky-200",
    title: "Mã xã không có trong danh mục",
    facility: "Quán Cơm tấm Sài Gòn",
    desc: "Phục vụ ăn uống tại chỗ, cơm phần, bún phở",
    code: "56101",
    workers: 6,
    revenue: "720 triệu",
    errorField: "admin",
    aiVerdict: {
      status: "error",
      title: "MÃ XÃ 999 KHÔNG HỢP LỆ",
      detected: "Mã xã 999 không thuộc danh mục hành chính cấp tỉnh.",
      suggest: "Đổi sang mã xã chuẩn theo địa bàn.",
      targetValue: "001"
    }
  }
];

interface OverviewDashboardProps {
  onNavigateTab: (tab: string) => void;
  mainDataLength: number;
  fileName: string;
  onFileUpload: (e: React.ChangeEvent<HTMLInputElement>) => void;
  onLoadSample: (data: any[]) => void;
}

export const OverviewDashboard: React.FC<OverviewDashboardProps> = ({
  onNavigateTab,
  mainDataLength,
  fileName,
  onFileUpload,
  onLoadSample
}) => {
  const [activeScenarioId, setActiveScenarioId] = useState<string>("vsic-mismatch");
  const [isScanning, setIsScanning] = useState<boolean>(false);
  const [scanCompleted, setScanCompleted] = useState<boolean>(true);
  const [fixedApplied, setFixedApplied] = useState<boolean>(false);

  const activeScenario = SCAN_SCENARIOS.find(s => s.id === activeScenarioId) || SCAN_SCENARIOS[0];

  const handleTriggerScan = () => {
    setIsScanning(true);
    setScanCompleted(false);
    setFixedApplied(false);
    setTimeout(() => {
      setIsScanning(false);
      setScanCompleted(true);
    }, 600);
  };

  const handleSelectScenario = (id: string) => {
    setActiveScenarioId(id);
    setFixedApplied(false);
    setIsScanning(true);
    setScanCompleted(false);
    setTimeout(() => {
      setIsScanning(false);
      setScanCompleted(true);
    }, 500);
  };

  return (
    <div className="space-y-4 animate-fade-in font-sans">
      {/* 1. KHỐI NẠP & QUẢN LÝ DỮ LIỆU THỐNG NHẤT (Nền xanh nhẹ mát mẻ) */}
      <div className="bg-white border border-sky-200/80 rounded-xl p-4 shadow-xs">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
          {/* Trạng thái dữ liệu */}
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-sky-50 border border-sky-200 flex items-center justify-center text-sky-700 shrink-0">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-800">DỮ LIỆU NGUỒN</span>
                {mainDataLength > 0 ? (
                  <span className="bg-emerald-50 text-emerald-700 border border-emerald-200 text-[10px] font-bold px-2 py-0.5 rounded-full">
                    {mainDataLength.toLocaleString()} dòng
                  </span>
                ) : (
                  <span className="bg-slate-100 text-slate-600 text-[10px] font-medium px-2 py-0.5 rounded-full">
                    Chưa có dữ liệu
                  </span>
                )}
              </div>
              <p className="text-xs text-slate-500 truncate max-w-md font-mono mt-0.5">
                {fileName || "Kéo thả hoặc chọn file Excel / CSV để kiểm tra"}
              </p>
            </div>
          </div>

          {/* Các nút thao tác nạp nhanh */}
          <div className="flex items-center gap-2 shrink-0 flex-wrap">
            {/* Nút tải tệp */}
            <label className="bg-[#144655] hover:bg-[#1a5567] text-white text-xs font-semibold px-3.5 py-2 rounded-lg transition-colors cursor-pointer flex items-center gap-1.5 shadow-xs">
              <UploadCloud className="w-4 h-4" /> Chọn tệp Excel / CSV
              <input
                type="file"
                accept=".xlsx,.xls,.csv,.txt"
                onChange={onFileUpload}
                className="hidden"
              />
            </label>

            {/* Nút nạp mẫu */}
            <button
              onClick={() => onLoadSample(SAMPLE_DATASET)}
              className="bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-200 text-xs font-semibold px-3.5 py-2 rounded-lg transition-colors cursor-pointer flex items-center gap-1.5"
            >
              <Play className="w-3.5 h-3.5 text-sky-700 fill-sky-700" /> Nạp mẫu 10 cơ sở
            </button>

            {/* Nút xem bảng dữ liệu nếu đã có dữ liệu */}
            {mainDataLength > 0 && (
              <button
                onClick={() => onNavigateTab("xemdulieu")}
                className="bg-[#fa9f4e] hover:bg-[#e88d3c] text-white text-xs font-semibold px-3.5 py-2 rounded-lg transition-colors cursor-pointer flex items-center gap-1 shadow-xs"
              >
                Xem bảng dữ liệu <ArrowRight className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>
      </div>

      {/* 2. KHỐI TRUNG TÂM: MÔ PHỎNG RÀ SOÁT LÔ-GÍCH & CÁC CÔNG CỤ CHÍNH (Gọn gàng trong 1 khối thống nhất) */}
      <div className="bg-white border border-sky-200/80 rounded-xl shadow-xs overflow-hidden">
        {/* Header khối */}
        <div className="bg-sky-50/60 border-b border-sky-100 px-4 py-2.5 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
          <div className="flex items-center gap-2">
            <Bot className="w-4 h-4 text-sky-700" />
            <span className="text-xs font-bold text-slate-800 uppercase tracking-wide">
              MÔ PHỎNG RÀ SOÁT LÔ-GÍCH
            </span>
            <span className="bg-sky-100 text-sky-800 text-[10px] font-bold px-2 py-0.5 rounded">
              AI RADAR
            </span>
          </div>

          <div className="flex items-center gap-2">
            {/* 3 nút tình huống nhỏ gọn */}
            <div className="flex items-center gap-1 bg-white p-0.5 rounded-lg border border-sky-200/80">
              {SCAN_SCENARIOS.map(sc => {
                const isSelected = activeScenarioId === sc.id;
                return (
                  <button
                    key={sc.id}
                    onClick={() => handleSelectScenario(sc.id)}
                    className={`px-2.5 py-1 text-xs rounded font-medium transition-all cursor-pointer ${
                      isSelected
                        ? "bg-sky-600 text-white font-bold shadow-xs"
                        : "text-slate-600 hover:text-slate-900 hover:bg-sky-50"
                    }`}
                  >
                    {sc.tag}
                  </button>
                );
              })}
            </div>

            {/* Nút Quét lại */}
            <button
              onClick={handleTriggerScan}
              disabled={isScanning}
              className="bg-sky-700 hover:bg-sky-800 text-white text-xs font-semibold px-3 py-1.5 rounded-lg shadow-xs flex items-center gap-1 cursor-pointer transition-colors"
            >
              {isScanning ? (
                <>
                  <RotateCcw className="w-3.5 h-3.5 animate-spin" /> Đang kiểm tra...
                </>
              ) : (
                <>
                  <Zap className="w-3.5 h-3.5 text-amber-300" /> Kiểm tra lại
                </>
              )}
            </button>
          </div>
        </div>

        {/* Bảng dòng dữ liệu kiểm tra */}
        <div className="p-4 space-y-3">
          <div className="overflow-x-auto rounded-lg border border-sky-100 bg-white">
            <table className="w-full text-left text-xs">
              <thead className="bg-sky-50/40 border-b border-sky-100 font-bold text-slate-600 text-[11px]">
                <tr>
                  <th className="px-3 py-2">CƠ SỞ</th>
                  <th className="px-3 py-2">MÔ TẢ THỰC TẾ</th>
                  <th className="px-3 py-2 text-center">MÃ KHAI BÁO</th>
                  <th className="px-3 py-2 text-center">LAO ĐỘNG</th>
                  <th className="px-3 py-2 text-right">DOANH THU</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-sky-50">
                <tr className="hover:bg-sky-50/30 transition-colors">
                  <td className="px-3 py-2.5 font-bold text-slate-800">{activeScenario.facility}</td>
                  <td className="px-3 py-2.5 text-slate-600 max-w-[280px]">
                    {activeScenario.desc}
                  </td>
                  
                  {/* Mã Ngành */}
                  <td className="px-3 py-2.5 text-center">
                    <span className={`inline-block font-mono font-bold px-2 py-0.5 rounded text-xs ${
                      activeScenario.errorField === "code" && scanCompleted && !fixedApplied
                        ? "bg-rose-100 text-rose-700 border border-rose-300"
                        : fixedApplied && activeScenario.errorField === "code"
                        ? "bg-emerald-100 text-emerald-800 border border-emerald-300"
                        : "bg-slate-100 text-slate-700"
                    }`}>
                      {fixedApplied && activeScenario.errorField === "code" 
                        ? activeScenario.aiVerdict.targetValue 
                        : activeScenario.code}
                    </span>
                  </td>

                  {/* Lao Động */}
                  <td className="px-3 py-2.5 text-center font-mono font-medium text-slate-700">
                    {activeScenario.workers} người
                  </td>

                  {/* Doanh Thu */}
                  <td className="px-3 py-2.5 text-right">
                    <span className={`inline-block font-mono font-bold px-2 py-0.5 rounded text-xs ${
                      activeScenario.errorField === "math" && scanCompleted && !fixedApplied
                        ? "bg-amber-100 text-amber-800 border border-amber-300"
                        : fixedApplied && activeScenario.errorField === "math"
                        ? "bg-emerald-100 text-emerald-800 border border-emerald-300"
                        : "text-slate-700"
                    }`}>
                      {fixedApplied && activeScenario.errorField === "math" 
                        ? activeScenario.aiVerdict.targetValue 
                        : activeScenario.revenue}
                    </span>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Dòng kết luận & hành động sửa nhanh */}
          {scanCompleted && (
            <div className={`p-3 rounded-lg border flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 ${
              fixedApplied 
                ? "bg-emerald-50 border-emerald-200 text-emerald-950" 
                : "bg-rose-50/70 border-rose-200 text-slate-800"
            }`}>
              <div className="flex items-center gap-2">
                {fixedApplied ? (
                  <span className="inline-flex items-center gap-1 bg-emerald-600 text-white text-[10px] font-bold px-2 py-0.5 rounded">
                    <Check className="w-3 h-3" /> ĐÃ SỬA
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1 bg-rose-600 text-white text-[10px] font-bold px-2 py-0.5 rounded">
                    <AlertTriangle className="w-3 h-3" /> {activeScenario.aiVerdict.title}
                  </span>
                )}
                <span className="text-xs text-slate-700">
                  {fixedApplied ? `Đã đồng bộ sang ${activeScenario.aiVerdict.targetValue}.` : activeScenario.aiVerdict.suggest}
                </span>
              </div>

              {!fixedApplied && activeScenario.aiVerdict.targetValue && (
                <button
                  onClick={() => setFixedApplied(true)}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3 py-1.5 rounded-lg flex items-center gap-1 cursor-pointer transition-colors shrink-0 self-start sm:self-auto shadow-xs"
                >
                  <CheckCircle2 className="w-3.5 h-3.5" /> Áp dụng mã "{activeScenario.aiVerdict.targetValue}"
                </button>
              )}
            </div>
          )}

          {/* Thanh truy cập nhanh các chức năng cốt lõi (Gọn gàng) */}
          <div className="pt-2 border-t border-sky-100 grid grid-cols-2 md:grid-cols-5 gap-2.5">
            <button
              onClick={() => onNavigateTab("xemdulieu")}
              className="flex items-center gap-2 p-2.5 rounded-lg bg-sky-50/50 hover:bg-sky-100/60 border border-sky-200/60 transition-colors text-left cursor-pointer group"
            >
              <FileSpreadsheet className="w-4 h-4 text-sky-700 shrink-0" />
              <div className="min-w-0">
                <div className="text-xs font-bold text-slate-800 truncate">1. Tải &amp; Gán cột</div>
                <div className="text-[10px] text-slate-500">Xem bảng dữ liệu</div>
              </div>
            </button>

            <button
              id="dash-shortcut-bctcxml"
              onClick={() => onNavigateTab("bctcxml")}
              className="flex items-center gap-2 p-2.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 transition-colors text-left cursor-pointer group shadow-2xs"
            >
              <FileSpreadsheet className="w-4 h-4 text-emerald-700 shrink-0" />
              <div className="min-w-0">
                <div className="text-xs font-bold text-emerald-950 truncate flex items-center gap-1">
                  BCTC XML <span className="bg-emerald-600 text-white text-[8px] px-1 rounded-sm">MỚI</span>
                </div>
                <div className="text-[10px] text-emerald-700 font-medium">Bóc tách HTKK thuế</div>
              </div>
            </button>

            <button
              onClick={() => onNavigateTab("chuanhoanganh")}
              className="flex items-center gap-2 p-2.5 rounded-lg bg-sky-50/50 hover:bg-sky-100/60 border border-sky-200/60 transition-colors text-left cursor-pointer group"
            >
              <Sparkles className="w-4 h-4 text-sky-700 shrink-0" />
              <div className="min-w-0">
                <div className="text-xs font-bold text-slate-800 truncate">2. Chuẩn hóa VSIC</div>
                <div className="text-[10px] text-slate-500">Đối chiếu mã ngành</div>
              </div>
            </button>

            <button
              onClick={() => onNavigateTab("kiemtralogic")}
              className="flex items-center gap-2 p-2.5 rounded-lg bg-sky-50/50 hover:bg-sky-100/60 border border-sky-200/60 transition-colors text-left cursor-pointer group"
            >
              <Scale className="w-4 h-4 text-sky-700 shrink-0" />
              <div className="min-w-0">
                <div className="text-xs font-bold text-slate-800 truncate">3. Kiểm tra logic</div>
                <div className="text-[10px] text-slate-500">Rà soát lỗi DN</div>
              </div>
            </button>

            <button
              onClick={() => onNavigateTab("tonghop")}
              className="flex items-center gap-2 p-2.5 rounded-lg bg-sky-50/50 hover:bg-sky-100/60 border border-sky-200/60 transition-colors text-left cursor-pointer group"
            >
              <BarChart3 className="w-4 h-4 text-sky-700 shrink-0" />
              <div className="min-w-0">
                <div className="text-xs font-bold text-slate-800 truncate">4. Báo cáo &amp; Xuất</div>
                <div className="text-[10px] text-slate-500">Tổng hợp kết quả</div>
              </div>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
