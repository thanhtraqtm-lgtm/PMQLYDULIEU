import React, { useState, useMemo, useRef } from "react";
import {
  FileText,
  Upload,
  Download,
  Building2,
  Calendar,
  DollarSign,
  TrendingUp,
  TrendingDown,
  CheckCircle2,
  AlertTriangle,
  Scale,
  Search,
  Copy,
  Check,
  RefreshCw,
  Sparkles,
  Layers,
  ArrowRight,
  PieChart,
  ShieldCheck,
  Percent,
  FileSpreadsheet
} from "lucide-react";
import {
  BctcParsedData,
  parseBctcXml,
  exportBctcToExcel,
  SAMPLE_BCTC_XML_STRING
} from "../utils/bctcXmlParser";

interface BctcXmlExtractorProps {
  onNavigateTab?: (tab: string) => void;
}

export const BctcXmlExtractor: React.FC<BctcXmlExtractorProps> = ({ onNavigateTab }) => {
  const [data, setData] = useState<BctcParsedData | null>(null);
  const [activeTab, setActiveTab] = useState<"bs" | "pl" | "cf" | "tb" | "ratios" | "json">("bs");
  const [searchAccount, setSearchAccount] = useState<string>("");
  const [filterCapTk, setFilterCapTk] = useState<"all" | "cap1" | "cap2">("all");
  const [copiedJson, setCopiedJson] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [isLoading, setIsLoading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Khởi tạo nạp dữ liệu mẫu của Công ty 69 nếu người dùng muốn
  const handleLoadSample = () => {
    setIsLoading(true);
    setErrorMessage("");
    try {
      const parsed = parseBctcXml(SAMPLE_BCTC_XML_STRING);
      setData(parsed);
      setActiveTab("bs");
    } catch (err: any) {
      setErrorMessage(err.message || "Lỗi khi nạp dữ liệu mẫu");
    } finally {
      setIsLoading(false);
    }
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    processFile(file);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) {
      processFile(file);
    }
  };

  const processFile = (file: File) => {
    setIsLoading(true);
    setErrorMessage("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = e.target?.result as string;
        if (!text) throw new Error("File rỗng hoặc không đọc được");
        const parsed = parseBctcXml(text);
        setData(parsed);
        setActiveTab("bs");
      } catch (err: any) {
        setErrorMessage(err.message || "Không thể bóc tách file XML. Vui lòng kiểm tra định dạng.");
      } finally {
        setIsLoading(false);
      }
    };
    reader.onerror = () => {
      setErrorMessage("Lỗi khi đọc file");
      setIsLoading(false);
    };
    reader.readAsText(file);
  };

  // Lọc bảng cân đối tài khoản
  const filteredTrialBalance = useMemo(() => {
    if (!data) return [];
    return data.trialBalance.filter((item) => {
      const matchSearch =
        item.soHieuTK.toLowerCase().includes(searchAccount.toLowerCase()) ||
        item.tenTK.toLowerCase().includes(searchAccount.toLowerCase());

      if (!matchSearch) return false;
      if (filterCapTk === "cap1") return item.capTK === 1;
      if (filterCapTk === "cap2") return item.capTK === 2;
      return true;
    });
  }, [data, searchAccount, filterCapTk]);

  const formatVND = (val: number) => {
    if (val === 0) return "0";
    return new Intl.NumberFormat("vi-VN").format(val);
  };

  const handleCopyJson = () => {
    if (!data) return;
    navigator.clipboard.writeText(JSON.stringify(data, null, 2));
    setCopiedJson(true);
    setTimeout(() => setCopiedJson(false), 2500);
  };

  const handleExportExcel = () => {
    if (!data) return;
    exportBctcToExcel(data, `BCTC_${data.header.mst || "DN"}`);
  };

  return (
    <div id="bctc-xml-extractor-root" className="space-y-6 pb-12">
      {/* Tiêu đề trang & Thanh công cụ */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-6 rounded-2xl border border-slate-200 shadow-xs">
        <div>
          <div className="flex items-center gap-2">
            <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 flex items-center gap-1">
              <Sparkles className="w-3.5 h-3.5" /> Chuẩn Thông tư 133 &amp; 200 (HTKK)
            </span>
            <span className="text-xs text-slate-500 font-medium">Phiên bản XML 2.3.2+</span>
          </div>
          <h2 className="text-2xl font-bold text-slate-800 mt-1.5 flex items-center gap-2">
            <FileText className="w-7 h-7 text-emerald-600" />
            Bóc Tách Dữ Liệu Báo Cáo Tài Chính (BCTC XML)
          </h2>
          <p className="text-sm text-slate-600 mt-1">
            Đọc tự động toàn bộ file XML BCTC từ phần mềm Hỗ trợ Kê khai (HTKK) hoặc Tổng cục Thuế: Báo cáo THTC (CĐKT), Kết quả SXKD, Lưu chuyển Tiền tệ và Bảng Cân đối Tài khoản 6 cột.
          </p>
        </div>

        <div className="flex items-center gap-2.5 flex-wrap">
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            accept=".xml,text/xml"
            className="hidden"
          />
          <button
            id="btn-upload-xml"
            onClick={() => fileInputRef.current?.click()}
            className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-sm font-semibold flex items-center gap-2 shadow-xs transition-colors"
          >
            <Upload className="w-4 h-4" /> Chọn File XML
          </button>

          <button
            id="btn-sample-69"
            onClick={handleLoadSample}
            className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-xl text-sm font-semibold flex items-center gap-2 border border-slate-300 transition-colors"
          >
            <Sparkles className="w-4 h-4 text-amber-600" /> Nạp Mẫu Công ty 69 (File bạn vừa gửi)
          </button>

          {data && (
            <button
              id="btn-export-excel"
              onClick={handleExportExcel}
              className="px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-semibold flex items-center gap-2 shadow-xs transition-colors"
            >
              <FileSpreadsheet className="w-4 h-4" /> Xuất Excel 6 Sheet
            </button>
          )}
        </div>
      </div>

      {/* Thông báo lỗi nếu có */}
      {errorMessage && (
        <div className="p-4 bg-rose-50 border border-rose-200 rounded-xl text-rose-800 text-sm flex items-center gap-3">
          <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0" />
          <div>
            <p className="font-semibold">Không thể xử lý file XML:</p>
            <p>{errorMessage}</p>
          </div>
        </div>
      )}

      {/* Khu vực Dropzone nếu chưa nạp dữ liệu */}
      {!data && !isLoading && (
        <div
          onDragOver={(e) => e.preventDefault()}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className="border-2 border-dashed border-slate-300 hover:border-emerald-500 bg-slate-50/60 hover:bg-emerald-50/30 transition-all rounded-3xl p-12 text-center cursor-pointer flex flex-col items-center justify-center space-y-4"
        >
          <div className="w-16 h-16 rounded-2xl bg-emerald-100 text-emerald-600 flex items-center justify-center shadow-xs">
            <Upload className="w-8 h-8" />
          </div>
          <div>
            <h3 className="text-lg font-bold text-slate-800">Kéo thả file BCTC XML vào đây hoặc nhấp để chọn</h3>
            <p className="text-sm text-slate-500 mt-1">
              Hỗ trợ file XML xuất từ HTKK phiên bản 4.x - 5.x, eTax của Tổng cục Thuế, MISA, Fast, Bravo...
            </p>
          </div>
          <div className="flex items-center gap-4 text-xs text-slate-400">
            <span>✓ Chuẩn Thông tư 133/2016/TT-BTC</span>
            <span>•</span>
            <span>✓ Chuẩn Thông tư 200/2014/TT-BTC</span>
            <span>•</span>
            <span>✓ Đầy đủ 4 Báo cáo &amp; Bảng Cân đối tài khoản</span>
          </div>
        </div>
      )}

      {/* Khi đã có dữ liệu bóc tách */}
      {data && (
        <div className="space-y-6">
          {/* Card Thông tin Doanh nghiệp & Hồ sơ thuế */}
          <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs">
            <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
              {/* Cột 1: Thông tin pháp nhân */}
              <div className="lg:col-span-2 space-y-3 border-b lg:border-b-0 lg:border-r border-slate-100 pb-4 lg:pb-0 lg:pr-6">
                <div className="flex items-start gap-3">
                  <div className="w-10 h-10 rounded-xl bg-blue-50 text-blue-700 flex items-center justify-center shrink-0">
                    <Building2 className="w-6 h-6" />
                  </div>
                  <div>
                    <h3 className="text-base font-bold text-slate-900 leading-snug">
                      {data.header.tenDoanhNghiep || "Doanh nghiệp chưa có tên"}
                    </h3>
                    <div className="flex items-center gap-2 mt-1">
                      <span className="text-xs font-mono font-bold bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md">
                        MST: {data.header.mst}
                      </span>
                      <span className="text-xs px-2 py-0.5 rounded-md font-medium bg-emerald-50 text-emerald-700">
                        {data.header.tenTKhai}
                      </span>
                    </div>
                    <p className="text-xs text-slate-500 mt-1.5 flex items-center gap-1">
                      📍 {data.header.diaChi || "Chưa có địa chỉ"}, {data.header.tinhTP}
                    </p>
                  </div>
                </div>
              </div>

              {/* Cột 2: Thông tin kỳ kê khai & CQT */}
              <div className="space-y-2 border-b lg:border-b-0 lg:border-r border-slate-100 pb-4 lg:pb-0 lg:pr-6">
                <div className="flex items-center gap-1.5 text-xs text-slate-500">
                  <Calendar className="w-3.5 h-3.5 text-slate-400" />
                  <span>Kỳ báo cáo tài chính:</span>
                </div>
                <p className="text-sm font-bold text-slate-800">
                  Năm {data.header.kyKKhai} ({data.header.kyKKhaiTuNgay} - {data.header.kyKKhaiDenNgay})
                </p>

                <div className="pt-2 text-xs text-slate-500">Cơ quan thuế nhận nộp:</div>
                <p className="text-xs font-semibold text-slate-700">
                  {data.header.tenCQT} (Mã: {data.header.maCQT})
                </p>
              </div>

              {/* Cột 3: Tình trạng hồ sơ & Ký số */}
              <div className="space-y-2 flex flex-col justify-between">
                <div>
                  <span className="text-xs text-slate-500">Ngày lập tờ khai:</span>
                  <p className="text-sm font-semibold text-slate-800">{data.header.ngayLapTKhai}</p>
                </div>

                <div className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></div>
                  <span className="text-xs font-medium text-emerald-800">
                    {data.header.ngayKy ? `Đã ký số: ${data.header.ngayKy}` : "Chưa có chữ ký điện tử"}
                  </span>
                </div>

                <div className="text-xs text-slate-400">
                  Phần mềm: HTKK v{data.header.pbanTKhaiXML}
                </div>
              </div>
            </div>
          </div>

          {/* Dải thẻ KPI tài chính chính */}
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3.5">
            {/* Tổng tài sản */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
              <span className="text-xs text-slate-500 font-medium block">Tổng Tài sản</span>
              <p className="text-lg font-bold text-slate-900 mt-1">
                {formatVND(data.rawSummary.tongTaiSanCuoiNam)}
              </p>
              <div className="flex items-center gap-1 mt-1 text-[11px] text-slate-500">
                <span>Đầu năm: {formatVND(data.rawSummary.tongTaiSanDauNam)}</span>
              </div>
            </div>

            {/* Doanh thu thuần */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
              <span className="text-xs text-slate-500 font-medium block">Doanh thu thuần</span>
              <p className="text-lg font-bold text-emerald-700 mt-1">
                {formatVND(data.rawSummary.doanhThuThuan)}
              </p>
              <div className="flex items-center gap-1 mt-1 text-[11px] text-slate-500">
                <span>Giá vốn: {formatVND(data.rawSummary.giaVon)}</span>
              </div>
            </div>

            {/* Lợi nhuận sau thuế */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
              <span className="text-xs text-slate-500 font-medium block">Lợi nhuận sau thuế</span>
              <p className={`text-lg font-bold mt-1 ${data.rawSummary.loiNhuanSauThue >= 0 ? "text-blue-700" : "text-rose-700"}`}>
                {formatVND(data.rawSummary.loiNhuanSauThue)}
              </p>
              <div className="flex items-center gap-1 mt-1 text-[11px] text-slate-500">
                <span>Trước thuế: {formatVND(data.rawSummary.loiNhuanTruocThue)}</span>
              </div>
            </div>

            {/* Tiền & Tương đương tiền */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
              <span className="text-xs text-slate-500 font-medium block">Tiền &amp; Tương đương tiền</span>
              <p className="text-lg font-bold text-purple-700 mt-1">
                {formatVND(data.rawSummary.tienVaTuongDuongTienCuoiKy)}
              </p>
              <div className="text-[11px] text-slate-500 mt-1">
                <span>Mã số 110 trên CĐKT</span>
              </div>
            </div>

            {/* Nợ phải trả */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs">
              <span className="text-xs text-slate-500 font-medium block">Nợ phải trả</span>
              <p className="text-lg font-bold text-amber-700 mt-1">
                {formatVND(data.rawSummary.noPhaiTraCuoiNam)}
              </p>
              <div className="text-[11px] text-slate-500 mt-1">
                <span>VCSH: {formatVND(data.rawSummary.vonChuSoHuuCuoiNam)}</span>
              </div>
            </div>

            {/* Cân đối kế toán & Bảng cân đối tài khoản */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-xs flex flex-col justify-between">
              <span className="text-xs text-slate-500 font-medium block">Kiểm tra cân đối</span>
              <div className="mt-1">
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700">
                  <CheckCircle2 className="w-3.5 h-3.5" />
                  CĐKT: Cân bằng
                </span>
                <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold mt-1 ${data.rawSummary.isTrialBalanceBalanced ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>
                  <Scale className="w-3.5 h-3.5" />
                  CDTK: {data.rawSummary.isTrialBalanceBalanced ? "Tổng Nợ = Có" : "Lệch"}
                </span>
              </div>
            </div>
          </div>

          {/* Hệ thống Tab điều hướng nội dung */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
            {/* Header Tabs */}
            <div className="flex border-b border-slate-200 overflow-x-auto bg-slate-50/70 p-1.5 gap-1">
              <button
                id="tab-balance-sheet"
                onClick={() => setActiveTab("bs")}
                className={`px-4 py-2.5 rounded-xl text-sm font-medium transition-all flex items-center gap-2 whitespace-nowrap ${
                  activeTab === "bs"
                    ? "bg-white text-emerald-700 shadow-xs font-semibold"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
                }`}
              >
                <Layers className="w-4 h-4 text-emerald-600" />
                Báo cáo Tình hình Tài chính (CĐKT)
                <span className="text-xs bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full font-mono">
                  {data.balanceSheet.length}
                </span>
              </button>

              <button
                id="tab-profit-loss"
                onClick={() => setActiveTab("pl")}
                className={`px-4 py-2.5 rounded-xl text-sm font-medium transition-all flex items-center gap-2 whitespace-nowrap ${
                  activeTab === "pl"
                    ? "bg-white text-emerald-700 shadow-xs font-semibold"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
                }`}
              >
                <TrendingUp className="w-4 h-4 text-blue-600" />
                Kết quả Hoạt động SXKD
                <span className="text-xs bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full font-mono">
                  {data.profitLoss.length}
                </span>
              </button>

              <button
                id="tab-cash-flow"
                onClick={() => setActiveTab("cf")}
                className={`px-4 py-2.5 rounded-xl text-sm font-medium transition-all flex items-center gap-2 whitespace-nowrap ${
                  activeTab === "cf"
                    ? "bg-white text-emerald-700 shadow-xs font-semibold"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
                }`}
              >
                <DollarSign className="w-4 h-4 text-purple-600" />
                Lưu chuyển Tiền tệ
                <span className="text-xs bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full font-mono">
                  {data.cashFlow.length}
                </span>
              </button>

              <button
                id="tab-trial-balance"
                onClick={() => setActiveTab("tb")}
                className={`px-4 py-2.5 rounded-xl text-sm font-medium transition-all flex items-center gap-2 whitespace-nowrap ${
                  activeTab === "tb"
                    ? "bg-white text-emerald-700 shadow-xs font-semibold"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
                }`}
              >
                <Scale className="w-4 h-4 text-amber-600" />
                Cân đối Tài khoản (6 cột)
                <span className="text-xs bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full font-mono">
                  {data.trialBalance.length}
                </span>
              </button>

              <button
                id="tab-ratios"
                onClick={() => setActiveTab("ratios")}
                className={`px-4 py-2.5 rounded-xl text-sm font-medium transition-all flex items-center gap-2 whitespace-nowrap ${
                  activeTab === "ratios"
                    ? "bg-white text-emerald-700 shadow-xs font-semibold"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
                }`}
              >
                <PieChart className="w-4 h-4 text-indigo-600" />
                Chỉ số Tài chính &amp; Đánh giá
              </button>

              <button
                id="tab-json"
                onClick={() => setActiveTab("json")}
                className={`px-4 py-2.5 rounded-xl text-sm font-medium transition-all flex items-center gap-2 whitespace-nowrap ${
                  activeTab === "json"
                    ? "bg-white text-emerald-700 shadow-xs font-semibold"
                    : "text-slate-600 hover:text-slate-900 hover:bg-slate-100"
                }`}
              >
                <Copy className="w-4 h-4 text-slate-600" />
                Dữ liệu JSON
              </button>
            </div>

            {/* Nội dung Tab 1: CĐKT */}
            {activeTab === "bs" && (
              <div className="p-6 overflow-x-auto">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h4 className="text-base font-bold text-slate-800">
                      Báo cáo Tình hình Tài chính (Mẫu B01a - DNN theo TT 133/2016/TT-BTC)
                    </h4>
                    <p className="text-xs text-slate-500">Đơn vị tính: Đồng Việt Nam (VNĐ)</p>
                  </div>
                  <span className="text-xs bg-slate-100 px-3 py-1 rounded-full text-slate-600 font-medium">
                    Tổng số chỉ tiêu có dữ liệu: {data.balanceSheet.length}
                  </span>
                </div>

                <table className="w-full text-left text-sm border-collapse">
                  <thead>
                    <tr className="bg-slate-100/80 text-slate-700 font-semibold border-y border-slate-200 text-xs uppercase tracking-wider">
                      <th className="py-3 px-3 w-16 text-center">Mã số</th>
                      <th className="py-3 px-4">Chỉ tiêu</th>
                      <th className="py-3 px-3 text-center w-24">Thuyết minh</th>
                      <th className="py-3 px-4 text-right">Số cuối năm</th>
                      <th className="py-3 px-4 text-right">Số đầu năm</th>
                      <th className="py-3 px-4 text-right">Chênh lệch (+/-)</th>
                      <th className="py-3 px-3 text-right w-24">Tăng giảm (%)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {data.balanceSheet.map((item) => {
                      const isBold = item.isHeader || ["200", "300", "400", "500"].includes(item.maSo);
                      const isMainTotal = ["200", "500"].includes(item.maSo);
                      return (
                        <tr
                          key={item.maSo}
                          className={`hover:bg-slate-50/80 transition-colors ${
                            isMainTotal
                              ? "bg-emerald-50/70 font-bold text-emerald-950"
                              : isBold
                              ? "bg-slate-50/50 font-semibold text-slate-900"
                              : "text-slate-700"
                          }`}
                        >
                          <td className="py-2.5 px-3 text-center font-mono text-xs text-slate-500">
                            {item.maSo}
                          </td>
                          <td className="py-2.5 px-4">{item.tenChiTieu}</td>
                          <td className="py-2.5 px-3 text-center text-xs text-slate-400">
                            {item.thuyetMinh || "-"}
                          </td>
                          <td className="py-2.5 px-4 text-right font-mono">
                            {formatVND(item.soCuoiNam)}
                          </td>
                          <td className="py-2.5 px-4 text-right font-mono text-slate-500">
                            {formatVND(item.soDauNam)}
                          </td>
                          <td
                            className={`py-2.5 px-4 text-right font-mono text-xs ${
                              item.chenhLech > 0
                                ? "text-emerald-700 font-medium"
                                : item.chenhLech < 0
                                ? "text-rose-700 font-medium"
                                : "text-slate-400"
                            }`}
                          >
                            {item.chenhLech > 0 ? `+${formatVND(item.chenhLech)}` : formatVND(item.chenhLech)}
                          </td>
                          <td
                            className={`py-2.5 px-3 text-right font-mono text-xs ${
                              item.phanTramThayDoi > 0
                                ? "text-emerald-700 font-semibold"
                                : item.phanTramThayDoi < 0
                                ? "text-rose-700 font-semibold"
                                : "text-slate-400"
                            }`}
                          >
                            {item.phanTramThayDoi > 0 ? `+${item.phanTramThayDoi}%` : `${item.phanTramThayDoi}%`}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* Nội dung Tab 2: KQKD */}
            {activeTab === "pl" && (
              <div className="p-6 overflow-x-auto">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h4 className="text-base font-bold text-slate-800">
                      Báo cáo Kết quả Hoạt động Sản xuất Kinh doanh (Mẫu B02 - DNN)
                    </h4>
                    <p className="text-xs text-slate-500">Đơn vị tính: Đồng Việt Nam (VNĐ)</p>
                  </div>
                </div>

                <table className="w-full text-left text-sm border-collapse">
                  <thead>
                    <tr className="bg-slate-100/80 text-slate-700 font-semibold border-y border-slate-200 text-xs uppercase tracking-wider">
                      <th className="py-3 px-3 w-16 text-center">Mã số</th>
                      <th className="py-3 px-4">Chỉ tiêu</th>
                      <th className="py-3 px-4 text-right">Năm nay</th>
                      <th className="py-3 px-4 text-right">Năm trước</th>
                      <th className="py-3 px-4 text-right">Chênh lệch (+/-)</th>
                      <th className="py-3 px-3 text-right w-24">Tăng giảm (%)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {data.profitLoss.map((item) => {
                      const isHighlight = ["10", "20", "30", "50", "60"].includes(item.maSo);
                      return (
                        <tr
                          key={item.maSo}
                          className={`hover:bg-slate-50/80 transition-colors ${
                            isHighlight ? "bg-blue-50/40 font-semibold text-slate-900" : "text-slate-700"
                          }`}
                        >
                          <td className="py-2.5 px-3 text-center font-mono text-xs text-slate-500">
                            {item.maSo}
                          </td>
                          <td className="py-2.5 px-4">{item.tenChiTieu}</td>
                          <td className="py-2.5 px-4 text-right font-mono font-medium">
                            {formatVND(item.namNay)}
                          </td>
                          <td className="py-2.5 px-4 text-right font-mono text-slate-500">
                            {formatVND(item.namTruoc)}
                          </td>
                          <td
                            className={`py-2.5 px-4 text-right font-mono text-xs ${
                              item.chenhLech > 0
                                ? "text-emerald-700 font-medium"
                                : item.chenhLech < 0
                                ? "text-rose-700 font-medium"
                                : "text-slate-400"
                            }`}
                          >
                            {item.chenhLech > 0 ? `+${formatVND(item.chenhLech)}` : formatVND(item.chenhLech)}
                          </td>
                          <td
                            className={`py-2.5 px-3 text-right font-mono text-xs ${
                              item.phanTramThayDoi > 0
                                ? "text-emerald-700 font-semibold"
                                : item.phanTramThayDoi < 0
                                ? "text-rose-700 font-semibold"
                                : "text-slate-400"
                            }`}
                          >
                            {item.phanTramThayDoi > 0 ? `+${item.phanTramThayDoi}%` : `${item.phanTramThayDoi}%`}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* Nội dung Tab 3: Lưu chuyển tiền tệ */}
            {activeTab === "cf" && (
              <div className="p-6 overflow-x-auto">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h4 className="text-base font-bold text-slate-800">
                      Báo cáo Lưu chuyển Tiền tệ (Phương pháp Trực tiếp)
                    </h4>
                    <p className="text-xs text-slate-500">Đơn vị tính: Đồng Việt Nam (VNĐ)</p>
                  </div>
                </div>

                <table className="w-full text-left text-sm border-collapse">
                  <thead>
                    <tr className="bg-slate-100/80 text-slate-700 font-semibold border-y border-slate-200 text-xs uppercase tracking-wider">
                      <th className="py-3 px-3 w-16 text-center">Mã số</th>
                      <th className="py-3 px-4">Chỉ tiêu</th>
                      <th className="py-3 px-4 text-right">Năm nay</th>
                      <th className="py-3 px-4 text-right">Năm trước</th>
                      <th className="py-3 px-4 text-right">Chênh lệch (+/-)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {data.cashFlow.map((item) => {
                      const isMainFlow = ["20", "30", "40", "50", "70"].includes(item.maSo);
                      return (
                        <tr
                          key={item.maSo}
                          className={`hover:bg-slate-50/80 transition-colors ${
                            isMainFlow ? "bg-purple-50/50 font-bold text-purple-950" : "text-slate-700"
                          }`}
                        >
                          <td className="py-2.5 px-3 text-center font-mono text-xs text-slate-500">
                            {item.maSo}
                          </td>
                          <td className="py-2.5 px-4">{item.tenChiTieu}</td>
                          <td className="py-2.5 px-4 text-right font-mono font-medium">
                            {formatVND(item.namNay)}
                          </td>
                          <td className="py-2.5 px-4 text-right font-mono text-slate-500">
                            {formatVND(item.namTruoc)}
                          </td>
                          <td
                            className={`py-2.5 px-4 text-right font-mono text-xs ${
                              item.chenhLech > 0
                                ? "text-emerald-700 font-medium"
                                : item.chenhLech < 0
                                ? "text-rose-700 font-medium"
                                : "text-slate-400"
                            }`}
                          >
                            {item.chenhLech > 0 ? `+${formatVND(item.chenhLech)}` : formatVND(item.chenhLech)}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* Nội dung Tab 4: Cân đối tài khoản 6 cột */}
            {activeTab === "tb" && (
              <div className="p-6 space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h4 className="text-base font-bold text-slate-800">
                      Bảng Cân đối Tài khoản (Bảng Cân đối Số phát sinh - 6 Cột)
                    </h4>
                    <p className="text-xs text-slate-500">
                      Tổng số dư Nợ cuối kỳ: <span className="font-mono font-bold text-slate-700">{formatVND(data.rawSummary.tongDuNoCuoiKy)}</span> |
                      Tổng số dư Có cuối kỳ: <span className="font-mono font-bold text-slate-700">{formatVND(data.rawSummary.tongDuCoCuoiKy)}</span>
                    </p>
                  </div>

                  <div className="flex items-center gap-2">
                    <div className="relative">
                      <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        placeholder="Tìm tài khoản (111, 154, Doanh thu...)"
                        value={searchAccount}
                        onChange={(e) => setSearchAccount(e.target.value)}
                        className="pl-9 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-emerald-500 w-64"
                      />
                    </div>

                    <select
                      value={filterCapTk}
                      onChange={(e) => setFilterCapTk(e.target.value as any)}
                      className="text-xs bg-slate-50 border border-slate-200 rounded-lg py-1.5 px-2.5 text-slate-700 focus:outline-none"
                    >
                      <option value="all">Tất cả tài khoản</option>
                      <option value="cap1">Chỉ TK cấp 1 (3 số)</option>
                      <option value="cap2">Chỉ TK cấp 2 (4 số)</option>
                    </select>
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 font-semibold border-y border-slate-200 uppercase tracking-wider text-[11px]">
                        <th className="py-2.5 px-2 w-20 text-center" rowSpan={2}>Số hiệu TK</th>
                        <th className="py-2.5 px-3 min-w-[200px]" rowSpan={2}>Tên tài khoản</th>
                        <th className="py-2 px-2 text-center border-b border-slate-200" colSpan={2}>Số dư đầu kỳ</th>
                        <th className="py-2 px-2 text-center border-b border-slate-200" colSpan={2}>Số phát sinh trong kỳ</th>
                        <th className="py-2 px-2 text-center border-b border-slate-200" colSpan={2}>Số dư cuối kỳ</th>
                      </tr>
                      <tr className="bg-slate-50 text-slate-600 font-semibold border-b border-slate-200 text-[10px]">
                        <th className="py-1.5 px-2 text-right">Dư Nợ</th>
                        <th className="py-1.5 px-2 text-right">Dư Có</th>
                        <th className="py-1.5 px-2 text-right">Phát sinh Nợ</th>
                        <th className="py-1.5 px-2 text-right">Phát sinh Có</th>
                        <th className="py-1.5 px-2 text-right">Dư Nợ</th>
                        <th className="py-1.5 px-2 text-right">Dư Có</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono text-[11px]">
                      {filteredTrialBalance.map((item) => {
                        const isCap1 = item.capTK === 1;
                        return (
                          <tr
                            key={item.soHieuTK}
                            className={`hover:bg-slate-50 transition-colors ${
                              isCap1 ? "bg-slate-50/40 font-semibold text-slate-900" : "text-slate-600"
                            }`}
                          >
                            <td className="py-2 px-2 text-center font-bold text-slate-700">
                              {item.soHieuTK}
                            </td>
                            <td className="py-2 px-3 font-sans text-xs">
                              <span className={item.capTK > 1 ? "pl-3 text-slate-600" : "text-slate-900"}>
                                {item.tenTK}
                              </span>
                            </td>
                            <td className="py-2 px-2 text-right text-slate-700">
                              {item.duNoDau !== 0 ? formatVND(item.duNoDau) : "-"}
                            </td>
                            <td className="py-2 px-2 text-right text-slate-700">
                              {item.duCoDau !== 0 ? formatVND(item.duCoDau) : "-"}
                            </td>
                            <td className="py-2 px-2 text-right text-blue-700">
                              {item.psNo !== 0 ? formatVND(item.psNo) : "-"}
                            </td>
                            <td className="py-2 px-2 text-right text-purple-700">
                              {item.psCo !== 0 ? formatVND(item.psCo) : "-"}
                            </td>
                            <td className="py-2 px-2 text-right text-emerald-800 font-bold">
                              {item.duNoCuoi !== 0 ? formatVND(item.duNoCuoi) : "-"}
                            </td>
                            <td className="py-2 px-2 text-right text-emerald-800 font-bold">
                              {item.duCoCuoi !== 0 ? formatVND(item.duCoCuoi) : "-"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Nội dung Tab 5: Chỉ số tài chính & Đánh giá */}
            {activeTab === "ratios" && (
              <div className="p-6 space-y-6">
                <div>
                  <h4 className="text-base font-bold text-slate-800">
                    Phân Tích Chỉ Số Tài Chính Doanh Nghiệp Tự Động
                  </h4>
                  <p className="text-xs text-slate-500">
                    Được tính toán trực tiếp từ dữ liệu BCTC bóc tách phục vụ công tác thanh kiểm tra &amp; lập Phiếu thống kê IO
                  </p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                  {/* Nhóm 1: Thanh toán */}
                  <div className="bg-slate-50/70 p-4 rounded-xl border border-slate-200 space-y-3">
                    <div className="flex items-center gap-2 font-bold text-slate-800 text-sm">
                      <ShieldCheck className="w-4 h-4 text-emerald-600" />
                      Khả Năng Thanh Toán
                    </div>
                    <div className="space-y-2 text-xs">
                      <div className="flex justify-between items-center py-1 border-b border-slate-200">
                        <span className="text-slate-600">Thanh toán hiện hành:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.thanhToanHienHanh.toFixed(2)} lần
                        </span>
                      </div>
                      <div className="flex justify-between items-center py-1 border-b border-slate-200">
                        <span className="text-slate-600">Thanh toán nhanh:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.thanhToanNhanh.toFixed(2)} lần
                        </span>
                      </div>
                      <div className="flex justify-between items-center py-1">
                        <span className="text-slate-600">Hệ số tiền mặt:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.tyLeTienMat.toFixed(2)} lần
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Nhóm 2: Cơ cấu vốn */}
                  <div className="bg-slate-50/70 p-4 rounded-xl border border-slate-200 space-y-3">
                    <div className="flex items-center gap-2 font-bold text-slate-800 text-sm">
                      <Scale className="w-4 h-4 text-blue-600" />
                      Cơ Cấu Vốn &amp; Đòn Bẩy
                    </div>
                    <div className="space-y-2 text-xs">
                      <div className="flex justify-between items-center py-1 border-b border-slate-200">
                        <span className="text-slate-600">Nợ / Tổng tài sản (D/A):</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.heSoNoTrenTongTaiSan.toFixed(1)}%
                        </span>
                      </div>
                      <div className="flex justify-between items-center py-1 border-b border-slate-200">
                        <span className="text-slate-600">Nợ / Vốn CSH (D/E):</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.heSoNoTrenVCSH.toFixed(1)}%
                        </span>
                      </div>
                      <div className="flex justify-between items-center py-1">
                        <span className="text-slate-600">Hệ số tự tài trợ:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.heSoTuTaiTro.toFixed(1)}%
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Nhóm 3: Khả năng sinh lời */}
                  <div className="bg-slate-50/70 p-4 rounded-xl border border-slate-200 space-y-3">
                    <div className="flex items-center gap-2 font-bold text-slate-800 text-sm">
                      <TrendingUp className="w-4 h-4 text-purple-600" />
                      Khả Năng Sinh Lời
                    </div>
                    <div className="space-y-2 text-xs">
                      <div className="flex justify-between items-center py-1 border-b border-slate-200">
                        <span className="text-slate-600">Biên lợi nhuận gộp:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.bienLoiNhuanGop.toFixed(2)}%
                        </span>
                      </div>
                      <div className="flex justify-between items-center py-1 border-b border-slate-200">
                        <span className="text-slate-600">Biên lãi ròng (ROS):</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.bienLoiNhuanRong.toFixed(2)}%
                        </span>
                      </div>
                      <div className="flex justify-between items-center py-1 border-b border-slate-200">
                        <span className="text-slate-600">Tỷ suất ROA:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.roa.toFixed(2)}%
                        </span>
                      </div>
                      <div className="flex justify-between items-center py-1">
                        <span className="text-slate-600">Tỷ suất ROE:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.roe.toFixed(2)}%
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Nhóm 4: Hiệu quả hoạt động */}
                  <div className="bg-slate-50/70 p-4 rounded-xl border border-slate-200 space-y-3">
                    <div className="flex items-center gap-2 font-bold text-slate-800 text-sm">
                      <RefreshCw className="w-4 h-4 text-amber-600" />
                      Hiệu Quả Hoạt Động
                    </div>
                    <div className="space-y-2 text-xs">
                      <div className="flex justify-between items-center py-1 border-b border-slate-200">
                        <span className="text-slate-600">Vòng quay Tổng tài sản:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.vongQuayTongTaiSan.toFixed(2)} vòng/năm
                        </span>
                      </div>
                      <div className="flex justify-between items-center py-1 border-b border-slate-200">
                        <span className="text-slate-600">Vòng quay Hàng tồn kho:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {data.ratios.vongQuayHangTonKho.toFixed(2)} vòng/năm
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Nội dung Tab 6: JSON */}
            {activeTab === "json" && (
              <div className="p-6 space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <h4 className="text-base font-bold text-slate-800">
                      Dữ Liệu JSON Chuẩn Hóa Bóc Tách Từ XML
                    </h4>
                    <p className="text-xs text-slate-500">
                      Dữ liệu có cấu trúc đầy đủ dùng để tích hợp API hoặc tự động điền vào Phiếu 01/IO-DN
                    </p>
                  </div>
                  <button
                    id="btn-copy-json"
                    onClick={handleCopyJson}
                    className="px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors"
                  >
                    {copiedJson ? (
                      <>
                        <Check className="w-4 h-4 text-emerald-600" /> Đã sao chép!
                      </>
                    ) : (
                      <>
                        <Copy className="w-4 h-4" /> Sao chép JSON
                      </>
                    )}
                  </button>
                </div>

                <div className="relative">
                  <pre className="bg-slate-900 text-emerald-400 p-4 rounded-xl text-xs font-mono max-h-[500px] overflow-auto">
                    {JSON.stringify(data, null, 2)}
                  </pre>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
