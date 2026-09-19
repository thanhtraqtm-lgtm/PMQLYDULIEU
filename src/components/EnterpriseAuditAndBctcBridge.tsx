import React, { useState, useMemo, useRef } from "react";
import {
  FileText,
  Upload,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Download,
  Eye,
  RefreshCw,
  Sparkles,
  Search,
  Building2,
  Calculator,
  ShieldAlert,
  ArrowRight,
  FileSpreadsheet,
  ChevronRight,
  ChevronDown,
  Info,
  Check,
  Edit3
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  RawBctcData,
  Phieu01IODNForm,
  AuditViolation,
  IndustrySectorProfile,
  INDUSTRY_PROFILES,
  detectIndustrySector,
  parseBctcXml,
  parseBctcPdfText,
  convertBctcToPhieu01IODN,
  auditPhieu01IODN,
  parsePhieu01Excel,
  parseBctcExcel,
  detectExcelFileType,
  createSampleDtvPhieu01
} from "../utils/bctcParser";
import {
  exportPhieu01IODNToExcel,
  exportBatchAuditExcel,
  generateSampleBctcWorkbook,
  exportSampleBctcExcelFile
} from "../utils/bctcExporter";
import { SAMPLE_BCTC_XML_STRING } from "../utils/bctcXmlParser";

interface EnterpriseAuditAndBctcBridgeProps {
  mainData?: any[];
  columns?: string[];
  fileName?: string;
  onNavigateTab?: (tab: string) => void;
}

export const EnterpriseAuditAndBctcBridge: React.FC<EnterpriseAuditAndBctcBridgeProps> = ({
  mainData = [],
  columns = [],
  fileName = "",
  onNavigateTab
}) => {
  const [activeSubTab, setActiveSubTab] = useState<"bctc_intake" | "batch_audit">("bctc_intake");

  // File BCTC state
  const [uploadedBctc, setUploadedBctc] = useState<RawBctcData | null>(null);
  const [generatedForm, setGeneratedForm] = useState<Phieu01IODNForm | null>(null);
  const [violations, setViolations] = useState<AuditViolation[]>([]);
  const [isProcessingBctc, setIsProcessingBctc] = useState<boolean>(false);
  const [selectedSectorId, setSelectedSectorId] = useState<string>("xay_dung");
  const [selectedMainIo, setSelectedMainIo] = useState<string>("105");
  const [bctcInputStatus, setBctcInputStatus] = useState<string>("");

  const xmlFileInputRef = useRef<HTMLInputElement>(null);
  const pdfFileInputRef = useRef<HTMLInputElement>(null);
  const bctcExcelInputRef = useRef<HTMLInputElement>(null);
  const phieuExcelInputRef = useRef<HTMLInputElement>(null);
  const batchExcelInputRef = useRef<HTMLInputElement>(null);

  // Mẫu dữ liệu BCTC thử nghiệm (TT200 & TT133)
  const handleLoadSampleBctc = (type: "TT200" | "TT133") => {
    const is200 = type === "TT200";
    const sample: RawBctcData = {
      mst: is200 ? "0101234567" : "0908765432",
      tenDoanhNghiep: is200
        ? "CÔNG TY CỔ PHẦN CHẾ BIẾN NÔNG SẢN XUẤT KHẨU HƯNG YÊN"
        : "CÔNG TY TNHH CƠ KHÍ & THƯƠNG MẠI HỒNG NAM",
      diaChi: is200
        ? "Khu công nghiệp Phố Nối A, Xã Lạc Hồng, Huyện Văn Lâm, Tỉnh Hưng Yên"
        : "Số 88 Đường Nguyễn Trãi, Phường Lam Sơn, TP Hưng Yên, Tỉnh Hưng Yên",
      nam: 2025,
      thongTu: type,
      dtt_10: is200 ? 58600000000 : 8450000000,
      gvhb_11: is200 ? 46200000000 : 6200000000,
      cfbh_25: is200 ? 3200000000 : 540000000,
      cfql_26: is200 ? 2800000000 : 480000000,
      laiVay_23: is200 ? 1150000000 : 120000000,
      ln_thuan_30: is200 ? 6400000000 : 1230000000,
      ln_truoc_thue_50: is200 ? 6550000000 : 1250000000,
      tonKho_141: is200 ? 8200000000 : 1450000000,
      cfsxdd_154_dauKy: is200 ? 1200000000 : 350000000,
      cfsxdd_154_cuoiKy: is200 ? 1450000000 : 420000000,
      thanhPham_155_dauKy: is200 ? 2100000000 : 480000000,
      thanhPham_155_cuoiKy: is200 ? 2400000000 : 510000000,
      hangGuiBan_157_dauKy: is200 ? 450000000 : 80000000,
      hangGuiBan_157_cuoiKy: is200 ? 520000000 : 95000000,
      hangHoa_156_dauKy: 0,
      hangHoa_156_cuoiKy: 0,
      cf_nvl_621: is200 ? 28500000000 : 3800000000,
      cf_nc_622: is200 ? 8600000000 : 1250000000,
      cf_sxc_627: is200 ? 4800000000 : 650000000,
      cf_mtc_623: 0,
      khauHao_214: is200 ? 2450000000 : 320000000,
      thue_333: is200 ? 1850000000 : 240000000,
      tienLuong_334: is200 ? 8600000000 : 1250000000,
      bhxh_338: is200 ? 1950000000 : 280000000,
      vayNganHan_320: is200 ? 6500000000 : 950000000,
      vayDaiHan_338: is200 ? 4200000000 : 0,
      sourceFileName: `Mau_BCTC_${type}_2025.xml`,
      sourceType: "xml"
    };

    setUploadedBctc(sample);
    const detected = detectIndustrySector(sample.tenDoanhNghiep);
    setSelectedSectorId(detected.id);
    setSelectedMainIo(detected.defaultIoCode);
    const form = convertBctcToPhieu01IODN(sample, detected.defaultIoCode, detected.id);
    setGeneratedForm(form);
    const issues = auditPhieu01IODN(form);
    setViolations(issues);
    setBctcInputStatus(`Đã nạp BCTC Mẫu (${type}) -> Tự động nhận diện ngành: "${detected.name}" (Mã IO: ${detected.defaultIoCode}) và hoàn thiện bóc tách phiếu!`);
  };

  // Nạp trực tiếp dữ liệu XML thực tế của Công ty 69 bạn gửi
  const handleLoadCty69Xml = () => {
    setIsProcessingBctc(true);
    setBctcInputStatus("Đang bóc tách dữ liệu XML BCTC Công ty 69...");
    try {
      const parsed = parseBctcXml(SAMPLE_BCTC_XML_STRING, "BCTC_CTY69_2025.xml");
      setUploadedBctc(parsed);
      const detected = detectIndustrySector(parsed.tenDoanhNghiep);
      setSelectedSectorId(detected.id);
      setSelectedMainIo(detected.defaultIoCode);
      const form = convertBctcToPhieu01IODN(parsed, detected.defaultIoCode, detected.id);
      setGeneratedForm(form);
      const issues = auditPhieu01IODN(form);
      setViolations(issues);
      setBctcInputStatus(
        `Đã bóc tách XML Công ty 69! Tự động nhận diện ngành: "${detected.name}" (Mã IO: ${detected.defaultIoCode}). Bóc tách tổng giá vốn ${(parsed.gvhb_11 || 0).toLocaleString("vi-VN")} đ vào các mã IO chi tiết.`
      );
    } catch (err: any) {
      setBctcInputStatus(`Lỗi bóc tách XML: ${err.message || "Không hợp lệ"}`);
    } finally {
      setIsProcessingBctc(false);
    }
  };

  // Nạp file XML BCTC
  const handleXmlUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessingBctc(true);
    setBctcInputStatus("Đang đọc, kiểm tra lỗi và nhận diện ngành nghề từ XML BCTC...");

    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const text = evt.target?.result as string;
        const parsed = parseBctcXml(text, file.name);
        setUploadedBctc(parsed);

        // Nhận diện ngành nghề thông minh và bóc tách
        const detected = detectIndustrySector(parsed.tenDoanhNghiep);
        setSelectedSectorId(detected.id);
        setSelectedMainIo(detected.defaultIoCode);

        // Chuyển đổi và điền vào Mẫu 01/IO-DN
        const form = convertBctcToPhieu01IODN(parsed, detected.defaultIoCode, detected.id);
        setGeneratedForm(form);

        // Kiểm tra lỗi & cảnh báo
        const issues = auditPhieu01IODN(form);
        setViolations(issues);

        setBctcInputStatus(
          `Đã đọc thành công BCTC XML [${file.name}]! Tự động nhận diện ngành nghề: "${detected.name}" (Mã IO: ${detected.defaultIoCode}), bóc tách từ Tổng vào các mã IO và hoàn thiện phiếu 01/IO-DN.`
        );
      } catch (err: any) {
        setBctcInputStatus(`Lỗi khi đọc file XML: ${err.message || "Tệp XML không đúng định dạng BCTC"}`);
      } finally {
        setIsProcessingBctc(false);
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  // Nạp file PDF BCTC
  const handlePdfUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessingBctc(true);
    setBctcInputStatus("Đang trích xuất nội dung văn bản tệp PDF BCTC...");

    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const decoder = new TextDecoder("utf-8");
        const raw = decoder.decode(buffer);

        // Trích xuất văn bản từ PDF stream
        const textMatches = raw.match(/\(([^)]+)\)\s*Tj/g) || raw.match(/\[([^\]]+)\]\s*TJ/g);
        let extracted = "";
        if (textMatches && textMatches.length > 0) {
          extracted = textMatches
            .map(m => m.replace(/[\(\)\[\]]/g, "").replace(/Tj|TJ/g, "").trim())
            .join(" ");
        } else {
          extracted = raw;
        }

        const parsed = parseBctcPdfText(extracted, file.name);
        setUploadedBctc(parsed);

        const detected = detectIndustrySector(parsed.tenDoanhNghiep);
        setSelectedSectorId(detected.id);
        setSelectedMainIo(detected.defaultIoCode);

        const form = convertBctcToPhieu01IODN(parsed, detected.defaultIoCode, detected.id);
        setGeneratedForm(form);

        const issues = auditPhieu01IODN(form);
        setViolations(issues);

        setBctcInputStatus(
          `Đã phân tích BCTC PDF [${file.name}]: Nhận diện "${detected.name}", tự động bóc tách vào các mã IO và tạo phiếu.`
        );
      } catch (err: any) {
        setBctcInputStatus(`Lỗi phân tích file PDF: ${err.message}`);
      } finally {
        setIsProcessingBctc(false);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  // Hàm xử lý file Excel thống nhất: Tự động phân biệt BCTC Excel hay Phiếu 01 của ĐTV
  const processExcelFile = (file: File) => {
    setIsProcessingBctc(true);
    setBctcInputStatus(`Đang đọc và phân tích cấu trúc file Excel [${file.name}]...`);

    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(buffer, { type: "array" });
        const detectedType = detectExcelFileType(wb);

        if (detectedType === "phieu01") {
          // File Mẫu Phiếu 01 của Điều tra viên
          const result = parsePhieu01Excel(buffer, file.name);
          setGeneratedForm(result.form);

          if (result.form.thongTinDinhDanh.sectorId) {
            setSelectedSectorId(result.form.thongTinDinhDanh.sectorId);
          }
          if (result.form.thongTinDinhDanh.maIO) {
            setSelectedMainIo(result.form.thongTinDinhDanh.maIO);
          }

          const issues = auditPhieu01IODN(result.form);
          setViolations(issues);

          const errorCount = issues.filter(i => i.type === "error").length;
          const warnCount = issues.filter(i => i.type === "warning").length;

          setBctcInputStatus(
            `Đã nạp thành công Phiếu 01 của ĐTV [${file.name}]! Tên DN: "${result.form.thongTinDinhDanh.tenDoanhNghiep}" (MST: ${result.form.thongTinDinhDanh.maSoThue}). Phát hiện ${errorCount} lỗi bắt buộc sửa và ${warnCount} cảnh báo cần giải trình!`
          );
        } else {
          // File Báo cáo tài chính Excel (B01 CĐKT, B02 KQKD, F01 CĐTK...)
          const parsed = parseBctcExcel(buffer, file.name);
          setUploadedBctc(parsed);

          const detected = detectIndustrySector(parsed.tenDoanhNghiep);
          setSelectedSectorId(detected.id);
          setSelectedMainIo(detected.defaultIoCode);

          const form = convertBctcToPhieu01IODN(parsed, detected.defaultIoCode, detected.id);
          setGeneratedForm(form);

          const issues = auditPhieu01IODN(form);
          setViolations(issues);

          const dttFormatted = parsed.dtt_10 > 0 ? `${parsed.dtt_10.toLocaleString("vi-VN")} đ` : "Chưa bóc tách được";
          const gvhbFormatted = parsed.gvhb_11 > 0 ? `${parsed.gvhb_11.toLocaleString("vi-VN")} đ` : "Chưa bóc tách được";

          setBctcInputStatus(
            `Đã nạp thành công BCTC Excel [${file.name}]! Đơn vị: "${parsed.tenDoanhNghiep}" (MST: ${parsed.mst || "Không có"}). DTT: ${dttFormatted}, Giá vốn: ${gvhbFormatted}. Tự động nhận diện ngành "${detected.name}" (Mã IO: ${detected.defaultIoCode}), bóc tách sang Phiếu 01/IO-DN và kiểm tra ${issues.length} chỉ tiêu!`
          );
        }
      } catch (err: any) {
        setBctcInputStatus(`Lỗi khi nạp file Excel: ${err.message || "Định dạng không hợp lệ"}`);
      } finally {
        setIsProcessingBctc(false);
      }
    };
    reader.onerror = () => {
      setBctcInputStatus(`Lỗi khi đọc file Excel [${file.name}]`);
      setIsProcessingBctc(false);
    };
    reader.readAsArrayBuffer(file);
  };

  // Nạp file Excel Báo cáo tài chính (B01, B02, F01)
  const handleBctcExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    processExcelFile(file);
    e.target.value = "";
  };

  // Nạp file Excel Phiếu 01/IO-DN của Điều tra viên
  const handlePhieuExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    processExcelFile(file);
    e.target.value = "";
  };

  // Nạp mẫu file Excel BCTC chuẩn (B01, B02, F01)
  const handleLoadSampleBctcExcel = () => {
    setIsProcessingBctc(true);
    setBctcInputStatus("Đang đọc và phân tích cấu trúc dữ liệu Mẫu Excel BCTC TT200...");
    try {
      const wb = generateSampleBctcWorkbook();
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const parsed = parseBctcExcel(wbout, "BCTC_Mau_TT200_HungYen_2025.xlsx");
      setUploadedBctc(parsed);

      const detected = detectIndustrySector(parsed.tenDoanhNghiep);
      setSelectedSectorId(detected.id);
      setSelectedMainIo(detected.defaultIoCode);

      const form = convertBctcToPhieu01IODN(parsed, detected.defaultIoCode, detected.id);
      setGeneratedForm(form);

      const issues = auditPhieu01IODN(form);
      setViolations(issues);

      setBctcInputStatus(
        `Đã nạp thành công Mẫu Excel BCTC (TT200)! DN: "${parsed.tenDoanhNghiep}" (MST: ${parsed.mst}). DTT: ${parsed.dtt_10.toLocaleString("vi-VN")} đ, Giá vốn: ${parsed.gvhb_11.toLocaleString("vi-VN")} đ. Tự động nhận diện ngành "${detected.name}" (Mã IO: ${detected.defaultIoCode}), bóc tách sang Phiếu 01/IO-DN và kiểm tra ${issues.length} chỉ tiêu!`
      );
    } catch (err: any) {
      setBctcInputStatus(`Lỗi khi nạp mẫu BCTC Excel: ${err.message || "Lỗi không xác định"}`);
    } finally {
      setIsProcessingBctc(false);
    }
  };

  // Kéo thả file bất kỳ (.xlsx, .xls, .xml, .pdf) vào vùng nạp
  const handleDropAnyFile = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (!file) return;

    const lowerName = file.name.toLowerCase();
    if (lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls")) {
      processExcelFile(file);
    } else if (lowerName.endsWith(".xml")) {
      const dummyEvent = { target: { files: [file], value: "" } } as any;
      handleXmlUpload(dummyEvent);
    } else if (lowerName.endsWith(".pdf")) {
      const dummyEvent = { target: { files: [file], value: "" } } as any;
      handlePdfUpload(dummyEvent);
    } else {
      setBctcInputStatus(`Định dạng file không hỗ trợ: ${file.name}. Vui lòng nạp file .xlsx, .xls, .xml hoặc .pdf`);
    }
  };

  // Nạp mẫu thử nghiệm Phiếu Excel thực tế của ĐTV (1001275111) có các lỗi thực tế
  const handleLoadSampleDtvPhieu = () => {
    setIsProcessingBctc(true);
    try {
      const sample = createSampleDtvPhieu01();
      setGeneratedForm(sample);
      setSelectedSectorId(sample.thongTinDinhDanh.sectorId || "det_may");
      setSelectedMainIo(sample.thongTinDinhDanh.maIO || "045");

      const issues = auditPhieu01IODN(sample);
      setViolations(issues);

      const errorCount = issues.filter(i => i.type === "error").length;
      const warnCount = issues.filter(i => i.type === "warning").length;

      setBctcInputStatus(
        `Đã nạp dữ liệu Mẫu Phiếu Điều tra viên (Phieu_01_1001275111.xlsx)! Phát hiện ${errorCount} lỗi nghiêm trọng và ${warnCount} cảnh báo thực tế (Có gia công nhưng phí gia công = 0, Giá vốn vượt Doanh thu, Có lương nhưng BHXH = 0, Khấu hao = 0, Lãi vay > 30% DTT...).`
      );
    } catch (err: any) {
      setBctcInputStatus(`Lỗi nạp mẫu phiếu ĐTV: ${err.message}`);
    } finally {
      setIsProcessingBctc(false);
    }
  };

  // Thay đổi nhóm ngành nghề doanh nghiệp (khi người dùng muốn điều chỉnh ngành)
  const handleChangeSector = (sectorId: string) => {
    setSelectedSectorId(sectorId);
    const sector = INDUSTRY_PROFILES.find(p => p.id === sectorId);
    const ioCode = sector ? sector.defaultIoCode : selectedMainIo;
    setSelectedMainIo(ioCode);
    if (uploadedBctc) {
      const form = convertBctcToPhieu01IODN(uploadedBctc, ioCode, sectorId);
      setGeneratedForm(form);
      const issues = auditPhieu01IODN(form);
      setViolations(issues);
    }
  };

  // Cập nhật ngành chính IO
  const handleChangeMainIo = (newIo: string) => {
    setSelectedMainIo(newIo);
    if (uploadedBctc) {
      const form = convertBctcToPhieu01IODN(uploadedBctc, newIo, selectedSectorId);
      setGeneratedForm(form);
      const issues = auditPhieu01IODN(form);
      setViolations(issues);
    }
  };

  // Cập nhật trực tiếp số tiền bóc tách Câu 15 (Nguyên vật liệu)
  const handleUpdateCau15Item = (index: number, field: "c1_giaTri" | "c2_tyLeNK" | "c3_giaCong", val: number) => {
    setGeneratedForm(prev => {
      if (!prev) return null;
      const updatedList = [...prev.phan3.cau15_nvlChiTiet];
      updatedList[index] = { ...updatedList[index], [field]: val };
      const updatedForm = {
        ...prev,
        phan3: { ...prev.phan3, cau15_nvlChiTiet: updatedList }
      };
      setViolations(auditPhieu01IODN(updatedForm));
      return updatedForm;
    });
  };

  // Cập nhật trực tiếp số tiền bóc tách Câu 18 (Dịch vụ mua ngoài)
  const handleUpdateCau18Item = (index: number, field: "c1_giaTri" | "c2_tyLeNK", val: number) => {
    setGeneratedForm(prev => {
      if (!prev) return null;
      const updatedList = [...prev.phan3.cau18_muaNgoaiChiPhiKhac];
      updatedList[index] = { ...updatedList[index], [field]: val };
      const updatedForm = {
        ...prev,
        phan3: { ...prev.phan3, cau18_muaNgoaiChiPhiKhac: updatedList }
      };
      setViolations(auditPhieu01IODN(updatedForm));
      return updatedForm;
    });
  };

  // Xuất file Excel Mẫu Phiếu 01/IO-DN
  const handleExportFormExcel = () => {
    if (!generatedForm) {
      alert("Vui lòng nạp Báo cáo tài chính trước khi xuất phiếu!");
      return;
    }
    exportPhieu01IODNToExcel(generatedForm, violations);
  };

  // Quản lý kiểm tra hàng loạt danh sách Doanh nghiệp từ file Excel
  const [batchData, setBatchData] = useState<any[]>(mainData);
  const [batchFileName, setBatchFileName] = useState<string>(fileName || "Dữ liệu hiện hành");

  const batchAuditResults = useMemo(() => {
    if (!batchData || batchData.length === 0) return [];

    return batchData.map((row, idx) => {
      // Nhận diện các cột thông minh
      const mst = row["MST"] || row["MaSoThue"] || row["Mã số thuế"] || row["mst"] || `DN_${idx + 1}`;
      const tenDN =
        row["TenDN"] ||
        row["Ten_DN"] ||
        row["Tên doanh nghiệp"] ||
        row["Tên DN"] ||
        row["ten_doanh_nghiep"] ||
        `Doanh nghiệp ${idx + 1}`;
      const dtt = Number(
        String(row["DoanhThu"] || row["DTT"] || row["Doanh thu"] || row["DTT_C5"] || 0).replace(/,/g, "")
      );
      const gvhb = Number(
        String(row["GiaVon"] || row["GVHB"] || row["Giá vốn"] || 0).replace(/,/g, "")
      );
      const cfbh = Number(String(row["CFBH"] || row["Chi phí bán hàng"] || 0).replace(/,/g, ""));
      const cfql = Number(String(row["CFQL"] || row["Chi phí quản lý"] || 0).replace(/,/g, ""));
      const cfnvl = Number(String(row["CFNVL"] || row["Nguyên vật liệu"] || 0).replace(/,/g, ""));
      const cfnc = Number(String(row["CFNC"] || row["Nhân công"] || 0).replace(/,/g, ""));
      const cfsxc = Number(String(row["CFSXC"] || row["Sản xuất chung"] || 0).replace(/,/g, ""));
      const khauHao = Number(String(row["KhauHao"] || row["KH"] || 0).replace(/,/g, ""));
      const laiVay = Number(String(row["LaiVay"] || row["Lãi vay"] || 0).replace(/,/g, ""));
      const maIO = String(row["MaIO"] || row["Mã IO"] || row["IO"] || "105");

      // Giả lập form để chạy hàm audit
      const mockRaw: RawBctcData = {
        mst: String(mst),
        tenDoanhNghiep: String(tenDN),
        diaChi: String(row["DiaChi"] || row["Địa chỉ"] || "Hưng Yên"),
        nam: 2025,
        thongTu: "TT200",
        dtt_10: dtt,
        gvhb_11: gvhb,
        cfbh_25: cfbh,
        cfql_26: cfql,
        laiVay_23: laiVay,
        ln_thuan_30: dtt - gvhb - cfbh - cfql,
        ln_truoc_thue_50: dtt - gvhb - cfbh - cfql,
        tonKho_141: 0,
        cfsxdd_154_dauKy: 0,
        cfsxdd_154_cuoiKy: 0,
        thanhPham_155_dauKy: 0,
        thanhPham_155_cuoiKy: 0,
        hangGuiBan_157_dauKy: 0,
        hangGuiBan_157_cuoiKy: 0,
        hangHoa_156_dauKy: 0,
        hangHoa_156_cuoiKy: 0,
        cf_nvl_621: cfnvl,
        cf_nc_622: cfnc,
        cf_sxc_627: cfsxc,
        cf_mtc_623: 0,
        khauHao_214: khauHao,
        thue_333: 0,
        tienLuong_334: cfnc,
        bhxh_338: cfnc > 0 ? cfnc * 0.22 : 0,
        vayNganHan_320: 0,
        vayDaiHan_338: 0,
        sourceFileName: "Batch",
        sourceType: "xml"
      };

      const mockForm = convertBctcToPhieu01IODN(mockRaw, maIO);
      const rowViolations = auditPhieu01IODN(mockForm);
      const errors = rowViolations.filter(v => v.type === "error");
      const warnings = rowViolations.filter(v => v.type === "warning");

      return {
        stt: idx + 1,
        tenDN: String(tenDN),
        mst: String(mst),
        soLoi: errors.length,
        soCanhBao: warnings.length,
        trangThai: errors.length > 0 ? "CÓ LỖI (Đỏ)" : warnings.length > 0 ? "CÓ CẢNH BÁO" : "ĐẠT CHUẨN",
        dtt,
        ic: mockForm.tongHopIO.ic,
        go: mockForm.tongHopIO.go,
        tyLeIC_GO: mockForm.tongHopIO.tyLeIC_GO,
        violations: rowViolations,
        chiTietLoi: rowViolations.map(v => `[${v.cau}] ${v.tenLoi}`).join("; ")
      };
    });
  }, [batchData]);

  // Bộ lọc bảng hàng loạt
  const [filterMode, setFilterMode] = useState<"all" | "errors_only" | "warnings_only">("all");
  const [searchDn, setSearchDn] = useState<string>("");

  const filteredBatchResults = useMemo(() => {
    return batchAuditResults.filter(item => {
      if (filterMode === "errors_only" && item.soLoi === 0) return false;
      if (filterMode === "warnings_only" && item.soCanhBao === 0) return false;
      if (searchDn.trim()) {
        const q = searchDn.toLowerCase();
        return item.tenDN.toLowerCase().includes(q) || item.mst.toLowerCase().includes(q);
      }
      return true;
    });
  }, [batchAuditResults, filterMode, searchDn]);

  const batchStats = useMemo(() => {
    const total = batchAuditResults.length;
    const errorCount = batchAuditResults.filter(r => r.soLoi > 0).length;
    const warningCount = batchAuditResults.filter(r => r.soLoi === 0 && r.soCanhBao > 0).length;
    const validCount = batchAuditResults.filter(r => r.soLoi === 0 && r.soCanhBao === 0).length;
    return { total, errorCount, warningCount, validCount };
  }, [batchAuditResults]);

  // Nạp file Excel riêng cho kiểm tra hàng loạt
  const handleBatchFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const json: any[] = XLSX.utils.sheet_to_json(ws, { defval: "" });

        if (json.length > 0) {
          setBatchData(json);
          setBatchFileName(file.name);
        } else {
          alert("Tệp Excel không có dòng dữ liệu!");
        }
      } catch (err: any) {
        alert("Không thể đọc tệp Excel: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  return (
    <div className="bg-white border border-slate-300 shadow-sm font-sans text-slate-800 space-y-4 pb-12">
      {/* 1. HEADER MODULE KIỂM TRA LỖI DN & NẠP BCTC CHUẨN IO 2026 */}
      <div className="bg-gradient-to-r from-emerald-800 via-teal-900 to-sky-900 text-white p-4 sm:p-5">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div>
            <div className="flex items-center gap-2">
              <span className="bg-amber-400 text-slate-950 text-[10px] font-black px-2 py-0.5 tracking-wider uppercase">
                QĐ 1124/QĐ-CTK • IO 2026
              </span>
              <span className="text-xs text-emerald-200 font-semibold">
                Thông tư 200/2014 & Thông tư 133/2016/TT-BTC
              </span>
            </div>
            <h1 className="text-lg sm:text-xl font-bold uppercase tracking-wide mt-1 m-0">
              CÔNG CỤ KIỂM TRA LỖI DOANH NGHIỆP & NẠP BCTC ĐIỀN MẪU PHIẾU 01/IO-DN
            </h1>
            <p className="text-xs text-emerald-100/90 mt-1 max-w-3xl leading-relaxed">
              Tự động đối chiếu các chỉ tiêu BCTC (XML / PDF), tự động điền 0 vào các chỉ tiêu không có,
              quét toàn bộ lỗi và cảnh báo theo đúng Hướng dẫn nghiệp vụ IO-2026, xuất biểu giải trình cho ĐTV và tệp Excel chuẩn mẫu.
            </p>
          </div>

          {/* TAB CHUYỂN ĐỔI CHỨC NĂNG */}
          <div className="flex items-center bg-white/10 p-1 border border-white/20 self-start md:self-auto shrink-0">
            <button
              type="button"
              onClick={() => setActiveSubTab("bctc_intake")}
              className={`px-3 py-1.5 text-xs font-bold flex items-center gap-1.5 cursor-pointer border-0 transition-colors ${
                activeSubTab === "bctc_intake"
                  ? "bg-white text-emerald-950 shadow-xs"
                  : "text-white/80 hover:text-white"
              }`}
            >
              <FileText className="w-4 h-4 text-emerald-600" />
              <span>Nạp BCTC & Điền Mẫu 01/IO-DN</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveSubTab("batch_audit")}
              className={`px-3 py-1.5 text-xs font-bold flex items-center gap-1.5 cursor-pointer border-0 transition-colors ${
                activeSubTab === "batch_audit"
                  ? "bg-white text-emerald-950 shadow-xs"
                  : "text-white/80 hover:text-white"
              }`}
            >
              <ShieldAlert className="w-4 h-4 text-amber-500" />
              <span>Kiểm Tra Lỗi Danh Sách DN ({batchData.length})</span>
            </button>
          </div>
        </div>
      </div>

      {/* ======================================================== */}
      {/* CHỨC NĂNG 1: NẠP BCTC (XML / PDF) -> ĐIỀN MẪU 01/IO-DN */}
      {/* ======================================================== */}
      {activeSubTab === "bctc_intake" && (
        <div className="px-3 sm:px-6 space-y-4">
          {/* KHỐI NẠP FILE BCTC (EXCEL / XML / PDF) HOẶC EXCEL ĐTV */}
          <div
            onDragOver={e => e.preventDefault()}
            onDrop={handleDropAnyFile}
            className="bg-slate-50 border border-slate-300 p-3.5 sm:p-4 space-y-3"
          >
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-2 border-b border-slate-200 pb-2">
              <div>
                <h3 className="text-sm font-bold text-slate-900 m-0 flex items-center gap-2">
                  <Upload className="w-4 h-4 text-emerald-700" />
                  NẠP DỮ LIỆU: BÁO CÁO TÀI CHÍNH (EXCEL/XML/PDF) HOẶC FILE EXCEL PHIẾU 01 ĐTV
                </h3>
                <p className="text-xs text-slate-600 m-0">
                  Hỗ trợ toàn diện file Excel BCTC (B01 CĐKT, B02 KQKD, F01 CĐTK từ MISA, FAST, Bravo, HTKK), file XML Thuế, PDF scan hoặc file Excel Phiếu 01 của ĐTV. Kéo thả file trực tiếp vào đây!
                </p>
              </div>

              {/* Nút nạp mẫu thử nhanh & tải file mẫu */}
              <div className="flex items-center flex-wrap gap-1.5">
                <span className="text-[11px] text-slate-500 font-bold">Thử nhanh:</span>
                <button
                  type="button"
                  onClick={handleLoadSampleBctcExcel}
                  className="text-xs bg-emerald-100 hover:bg-emerald-200 text-emerald-950 border border-emerald-400 font-bold px-2.5 py-1 cursor-pointer shadow-xs flex items-center gap-1"
                  title="Nạp ngay dữ liệu mẫu file Excel BCTC TT200 (Cơ khí & chế tạo máy Hưng Yên) có đủ B01, B02, F01"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
                  Mẫu BCTC Excel (TT200)
                </button>
                <button
                  type="button"
                  onClick={exportSampleBctcExcelFile}
                  className="text-xs bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-bold px-2 py-1 cursor-pointer flex items-center gap-1"
                  title="Tải về máy file Excel BCTC chuẩn gồm 3 biểu B01-DN, B02-DN, F01-DN để tham khảo cấu trúc"
                >
                  <Download className="w-3 h-3 text-slate-500" />
                  Tải mẫu Excel BCTC
                </button>
                <button
                  type="button"
                  onClick={() => handleLoadSampleBctc("TT200")}
                  className="text-xs bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-300 font-bold px-2 py-1 cursor-pointer"
                >
                  Mẫu TT 200 (XML)
                </button>
                <button
                  type="button"
                  onClick={() => handleLoadSampleBctc("TT133")}
                  className="text-xs bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-300 font-bold px-2 py-1 cursor-pointer"
                >
                  Mẫu TT 133 (XML)
                </button>
                <button
                  type="button"
                  onClick={handleLoadCty69Xml}
                  className="text-xs bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 font-bold px-2 py-1 cursor-pointer"
                  title="Nạp tệp XML Công ty 69 đã giải mã"
                >
                  Mẫu Cty 69 (XML)
                </button>
                <button
                  type="button"
                  onClick={handleLoadSampleDtvPhieu}
                  className="text-xs bg-purple-100 hover:bg-purple-200 text-purple-950 border border-purple-400 font-bold px-2 py-1 cursor-pointer shadow-xs flex items-center gap-1"
                  title="Nạp mẫu file Excel Phiếu 01 thực tế của điều tra viên (1001275111) để xem ngay danh sách các lỗi điều tra viên hay mắc"
                >
                  <ShieldAlert className="w-3.5 h-3.5 text-purple-700" />
                  Mẫu Phiếu ĐTV Excel
                </button>
              </div>
            </div>

            {/* Các nút tải file: Lưới 5 ô với Nạp Excel BCTC nổi bật nhất */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
              {/* 1. TẢI FILE EXCEL BÁO CÁO TÀI CHÍNH (Nổi bật nhất) */}
              <div
                onClick={() => bctcExcelInputRef.current?.click()}
                className="border-2 border-dashed border-emerald-500 hover:border-emerald-600 bg-emerald-50/70 hover:bg-emerald-100/70 p-3 text-center cursor-pointer transition-colors shadow-xs"
              >
                <FileSpreadsheet className="w-6 h-6 text-emerald-700 mx-auto mb-1" />
                <span className="text-xs font-bold text-emerald-950 block">1. Nạp file Excel BCTC</span>
                <span className="text-[10px] text-emerald-900 font-semibold block mt-0.5">(.xlsx, .xls) B01, B02, F01</span>
                <span className="text-[9.5px] text-slate-500 block leading-tight mt-0.5">Từ MISA, FAST, Bravo, HTKK...</span>
                <input
                  type="file"
                  ref={bctcExcelInputRef}
                  onChange={handleBctcExcelUpload}
                  accept=".xlsx, .xls"
                  className="hidden"
                />
              </div>

              {/* 2. Tải XML BCTC */}
              <div
                onClick={() => xmlFileInputRef.current?.click()}
                className="border-2 border-dashed border-sky-300 hover:border-sky-500 bg-white hover:bg-sky-50/50 p-3 text-center cursor-pointer transition-colors"
              >
                <FileText className="w-6 h-6 text-sky-600 mx-auto mb-1" />
                <span className="text-xs font-bold text-sky-900 block">2. Nạp file XML BCTC</span>
                <span className="text-[10px] text-slate-500 block mt-0.5">File XML xuất từ HTKK, Thuế điện tử eTax</span>
                <input
                  type="file"
                  ref={xmlFileInputRef}
                  onChange={handleXmlUpload}
                  accept=".xml"
                  className="hidden"
                />
              </div>

              {/* 3. Tải PDF BCTC */}
              <div
                onClick={() => pdfFileInputRef.current?.click()}
                className="border-2 border-dashed border-rose-300 hover:border-rose-500 bg-white hover:bg-rose-50/50 p-3 text-center cursor-pointer transition-colors"
              >
                <FileText className="w-6 h-6 text-rose-600 mx-auto mb-1" />
                <span className="text-xs font-bold text-rose-900 block">3. Nạp file PDF BCTC</span>
                <span className="text-[10px] text-slate-500 block mt-0.5">BCTC PDF scan hoặc dạng văn bản</span>
                <input
                  type="file"
                  ref={pdfFileInputRef}
                  onChange={handlePdfUpload}
                  accept=".pdf"
                  className="hidden"
                />
              </div>

              {/* 4. Tải FILE EXCEL PHIẾU 01 CỦA ĐIỀU TRA VIÊN */}
              <div
                onClick={() => phieuExcelInputRef.current?.click()}
                className="border-2 border-dashed border-purple-400 hover:border-purple-600 bg-purple-50/40 hover:bg-purple-50 p-3 text-center cursor-pointer transition-colors"
              >
                <FileSpreadsheet className="w-6 h-6 text-purple-700 mx-auto mb-1" />
                <span className="text-xs font-bold text-purple-950 block">4. Nạp Excel Phiếu 01 ĐTV</span>
                <span className="text-[10px] text-purple-800 font-medium block mt-0.5">Kiểm tra lỗi điều tra viên kê khai</span>
                <input
                  type="file"
                  ref={phieuExcelInputRef}
                  onChange={handlePhieuExcelUpload}
                  accept=".xlsx, .xls"
                  className="hidden"
                />
              </div>

              {/* 5. Nhận diện và chọn Nhóm ngành nghề DN */}
              <div className="bg-white border border-slate-300 p-2.5 flex flex-col justify-between space-y-1.5">
                <div>
                  <label className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                    <Building2 className="w-3.5 h-3.5 text-sky-700" />
                    5. Ngành nghề & Bóc tách:
                  </label>
                  <select
                    value={selectedSectorId}
                    onChange={e => handleChangeSector(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-300 text-xs font-bold p-1 text-slate-900 focus:outline-none focus:border-sky-600 mt-1"
                  >
                    {INDUSTRY_PROFILES.map(prof => (
                      <option key={prof.id} value={prof.id}>
                        {prof.name} (Mã IO: {prof.defaultIoCode})
                      </option>
                    ))}
                  </select>
                </div>
                <div className="flex items-center justify-between gap-1 pt-1 border-t border-slate-100">
                  <span className="text-[11px] font-bold text-slate-600">Mã IO:</span>
                  <select
                    value={selectedMainIo}
                    onChange={e => handleChangeMainIo(e.target.value)}
                    className="bg-slate-50 border border-slate-300 text-xs font-bold px-1.5 py-0.5 text-slate-900 focus:outline-none focus:border-emerald-600"
                  >
                    <option value="105">105. Chế biến thực phẩm</option>
                    <option value="075">075. Kim loại, thép</option>
                    <option value="082">082. Kim loại đúc sẵn, cơ khí</option>
                    <option value="045">045. May mặc, trang phục</option>
                    <option value="123">123. Xây dựng công trình nhà</option>
                    <option value="124">124. Xây dựng cầu đường</option>
                    <option value="129">129. Bán buôn hàng hóa</option>
                    <option value="130">130. Bán lẻ hàng hóa</option>
                    <option value="133">133. Vận tải đường bộ</option>
                    <option value="145">145. CNTT & phần mềm</option>
                    <option value="099">099. Công nghiệp khác</option>
                  </select>
                </div>
              </div>
            </div>

            {bctcInputStatus && (
              <div className="bg-emerald-50 border border-emerald-300 text-emerald-900 text-xs p-2 font-medium flex items-center gap-1.5">
                <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
                <span>{bctcInputStatus}</span>
              </div>
            )}
          </div>

          {/* NẾU ĐÃ CÓ PHIẾU ĐƯỢC TẠO: HIỂN THỊ KẾT QUẢ VÀ KIỂM TRA LỖI */}
          {generatedForm && (
            <div className="space-y-4">
              {/* BẢNG TỔNG QUAN TÌNH TRẠNG LỖI & CẢNH BÁO */}
              <div
                className={`border p-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
                  violations.some(v => v.type === "error")
                    ? "bg-rose-50 border-rose-300 text-rose-950"
                    : violations.length > 0
                    ? "bg-amber-50 border-amber-300 text-amber-950"
                    : "bg-emerald-50 border-emerald-300 text-emerald-950"
                }`}
              >
                <div className="space-y-0.5">
                  <div className="flex items-center gap-2">
                    {violations.some(v => v.type === "error") ? (
                      <XCircle className="w-5 h-5 text-rose-600 shrink-0" />
                    ) : violations.length > 0 ? (
                      <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0" />
                    ) : (
                      <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
                    )}
                    <span className="font-bold text-sm">
                      KẾT QUẢ KIỂM TRA LỖI PHIẾU 01/IO-DN THEO HƯỚNG DẪN IO-2026:{" "}
                      {violations.filter(v => v.type === "error").length} LỖI BẮT BUỘC SỬA,{" "}
                      {violations.filter(v => v.type === "warning").length} CẢNH BÁO CẦN XÁC MINH
                    </span>
                  </div>
                  <p className="text-xs text-slate-600 m-0 pl-7">
                    Doanh nghiệp: <strong>{generatedForm.thongTinDinhDanh.tenDoanhNghiep}</strong> (MST:{" "}
                    <strong>{generatedForm.thongTinDinhDanh.maSoThue}</strong>) • Tỷ lệ IC/GO sơ bộ:{" "}
                    <strong>{(generatedForm.tongHopIO.tyLeIC_GO * 100).toFixed(1)}%</strong> ({generatedForm.tongHopIO.danhGiaIC})
                  </p>
                </div>

                <div className="flex items-center gap-2 pl-7 sm:pl-0">
                  <button
                    type="button"
                    onClick={handleExportFormExcel}
                    className="bg-[#28a745] hover:bg-[#218838] text-white font-bold text-xs px-4 py-2 flex items-center gap-1.5 cursor-pointer border-0 shadow-xs"
                    title="Xuất Mẫu Phiếu 01/IO-DN hoàn chỉnh kèm Biểu giải trình lỗi ra file Excel"
                  >
                    <Download className="w-4 h-4 text-white" />
                    <span>Xuất File Excel Phiếu 01/IO-DN</span>
                  </button>
                </div>
              </div>

              {/* DANH SÁCH CHI TIẾT LỖI VÀ CẢNH BÁO (BIỂU GIẢI TRÌNH ĐTV) */}
              {violations.length > 0 && (
                <div className="border border-slate-300 bg-white overflow-hidden">
                  <div className="bg-slate-100 px-3 py-2 border-b border-slate-200 flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-800 uppercase flex items-center gap-1.5">
                      <ShieldAlert className="w-4 h-4 text-rose-600" />
                      Biểu Giải Trình Lỗi & Cảnh Báo (Đề nghị Điều tra viên giải trình theo QĐ 1124)
                    </span>
                    <span className="text-[11px] text-slate-500 font-medium">({violations.length} nội dung cần chú ý)</span>
                  </div>
                  <div className="overflow-x-auto max-h-60">
                    <table className="w-full text-xs text-left border-collapse">
                      <thead className="bg-slate-50 text-slate-700 font-bold border-b border-slate-200 sticky top-0">
                        <tr>
                          <th className="py-1.5 px-2 text-center w-10">STT</th>
                          <th className="py-1.5 px-2 w-24">Vị trí câu</th>
                          <th className="py-1.5 px-2 w-24">Phân loại</th>
                          <th className="py-1.5 px-3">Nội dung lỗi & Quy định kiểm tra</th>
                          <th className="py-1.5 px-2 text-right w-36">Số liệu thực tế</th>
                          <th className="py-1.5 px-3">Hướng dẫn xử lý cho ĐTV</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {violations.map((v, vIdx) => (
                          <tr key={v.id} className={v.type === "error" ? "bg-rose-50/50" : "bg-amber-50/30"}>
                            <td className="py-1.5 px-2 text-center font-bold text-slate-500">{vIdx + 1}</td>
                            <td className="py-1.5 px-2 font-bold text-sky-800">{v.cau}</td>
                            <td className="py-1.5 px-2">
                              {v.type === "error" ? (
                                <span className="bg-rose-600 text-white font-bold text-[10px] px-1.5 py-0.2">
                                  LỖI
                                </span>
                              ) : (
                                <span className="bg-amber-500 text-slate-950 font-bold text-[10px] px-1.5 py-0.2">
                                  CẢNH BÁO
                                </span>
                              )}
                            </td>
                            <td className="py-1.5 px-3 text-slate-900 font-medium">{v.moTa}</td>
                            <td className="py-1.5 px-2 text-right font-mono font-bold text-rose-800">{v.giaTriThucTe}</td>
                            <td className="py-1.5 px-3 text-slate-600 italic">{v.huongDanXuLy}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* GIAO DIỆN TRỰC QUAN MẪU PHIẾU 01/IO-DN (CHỈ TIÊU KHÔNG CÓ TRÊN BCTC ĐÃ ĐƯỢC TỰ ĐỘNG ĐIỀN 0) */}
              <div className="border border-sky-400 bg-white">
                <div className="bg-[#286e42] text-white px-3 py-2 flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <FileSpreadsheet className="w-4 h-4 text-emerald-200" />
                    <span className="font-bold text-xs uppercase">
                      XEM TRƯỚC MẪU PHIẾU 01/IO-DN (TỰ ĐỘNG ĐIỀN CHỈ TIÊU TỪ BCTC • CÒN LẠI ĐIỀN SỐ 0)
                    </span>
                  </div>
                  <span className="text-[11px] text-emerald-100 italic">
                    (Có thể chỉnh sửa trực tiếp từng ô số liệu nếu cần bổ sung)
                  </span>
                </div>

                <div className="p-3 sm:p-4 space-y-4 max-h-[600px] overflow-y-auto text-xs">
                  {/* THÔNG TIN ĐỊNH DANH */}
                  <div className="border border-slate-200 p-3 bg-slate-50/50 space-y-2.5">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200 pb-1.5">
                      <span className="font-bold text-sky-900 block uppercase">
                        THÔNG TIN ĐỊNH DANH & NHẬN DIỆN NGÀNH NGHỀ DOANH NGHIỆP
                      </span>
                      <div className="flex items-center gap-2">
                        <span className="bg-sky-100 text-sky-900 px-2 py-0.5 font-bold text-[11px] border border-sky-300">
                          Mã C5: {generatedForm.thongTinDinhDanh.maNganhC5}
                        </span>
                        <span className="bg-emerald-100 text-emerald-900 px-2 py-0.5 font-bold text-[11px] border border-emerald-300">
                          Mã IO: {generatedForm.thongTinDinhDanh.maIO}
                        </span>
                      </div>
                    </div>

                    {/* Banner Nhận diện ngành nghề & Bóc tách */}
                    {generatedForm.thongTinDinhDanh.sectorExplanation && (
                      <div className="bg-sky-50 border border-sky-300 p-2 text-sky-950 flex items-start gap-2">
                        <Sparkles className="w-4 h-4 text-sky-600 shrink-0 mt-0.5" />
                        <div className="space-y-0.5">
                          <span className="font-bold block text-sky-900">
                            Ngành nhận diện: {generatedForm.thongTinDinhDanh.tenNganhChinh} (Mã IO: {generatedForm.thongTinDinhDanh.maIO})
                          </span>
                          <p className="text-[11px] text-sky-800 m-0 leading-relaxed">
                            {generatedForm.thongTinDinhDanh.sectorExplanation} Toàn bộ tổng chi phí NVL & dịch vụ mua ngoài từ BCTC được tự động bóc tách thành các mã IO chi tiết tương ứng ngành.
                          </p>
                        </div>
                      </div>
                    )}

                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                      <div>
                        <span className="text-slate-500 block text-[11px]">Tên doanh nghiệp:</span>
                        <input
                          type="text"
                          value={generatedForm.thongTinDinhDanh.tenDoanhNghiep}
                          onChange={e => {
                            const val = e.target.value;
                            setGeneratedForm(prev => prev && { ...prev, thongTinDinhDanh: { ...prev.thongTinDinhDanh, tenDoanhNghiep: val } });
                          }}
                          className="w-full bg-white border border-slate-300 px-2 py-1 font-bold text-slate-800"
                        />
                      </div>
                      <div>
                        <span className="text-slate-500 block text-[11px]">Mã số thuế:</span>
                        <input
                          type="text"
                          value={generatedForm.thongTinDinhDanh.maSoThue}
                          onChange={e => {
                            const val = e.target.value;
                            setGeneratedForm(prev => prev && { ...prev, thongTinDinhDanh: { ...prev.thongTinDinhDanh, maSoThue: val } });
                          }}
                          className="w-full bg-white border border-slate-300 px-2 py-1 font-mono font-bold text-slate-800"
                        />
                      </div>
                      <div>
                        <span className="text-slate-500 block text-[11px]">Địa chỉ:</span>
                        <input
                          type="text"
                          value={generatedForm.thongTinDinhDanh.diaChi}
                          onChange={e => {
                            const val = e.target.value;
                            setGeneratedForm(prev => prev && { ...prev, thongTinDinhDanh: { ...prev.thongTinDinhDanh, diaChi: val } });
                          }}
                          className="w-full bg-white border border-slate-300 px-2 py-1 text-slate-800"
                        />
                      </div>
                    </div>
                  </div>

                  {/* PHẦN I: KẾT QUẢ KINH DOANH */}
                  <div className="border border-slate-200 p-3 bg-white space-y-2">
                    <span className="font-bold text-sky-900 block border-b border-slate-200 pb-1 uppercase">
                      PHẦN I: KẾT QUẢ HOẠT ĐỘNG SẢN XUẤT KINH DOANH (CÂU 5, 6, 7, 8, 9)
                    </span>

                    {/* Câu 5 */}
                    <div className="flex items-center justify-between bg-sky-50/60 p-2 border border-sky-200">
                      <span className="font-bold text-slate-800">
                        5. Tổng doanh thu thuần hoạt động SXKD của DN năm 2025 (đồng):
                      </span>
                      <input
                        type="number"
                        value={generatedForm.phan1.cau5_tongDoanhThuThuan}
                        onChange={e => {
                          const val = Number(e.target.value);
                          setGeneratedForm(prev => prev && { ...prev, phan1: { ...prev.phan1, cau5_tongDoanhThuThuan: val } });
                        }}
                        className="bg-white border border-sky-300 px-2 py-1 text-right font-mono font-bold text-slate-900 w-48"
                      />
                    </div>

                    {/* Bảng Câu 6 */}
                    <div className="border border-slate-200 mt-2">
                      <div className="bg-slate-100 p-1.5 font-bold text-slate-800 border-b border-slate-200">
                        6. Doanh thu thuần chia theo sản phẩm SXKD của DN năm 2025
                      </div>
                      <table className="w-full text-left border-collapse">
                        <thead className="bg-slate-50 text-[11px] text-slate-600 border-b border-slate-200">
                          <tr>
                            <th className="p-1.5">Tên ngành sản phẩm</th>
                            <th className="p-1.5 w-20">Mã IO</th>
                            <th className="p-1.5 text-right w-44">Doanh thu thuần (Cột 1)</th>
                            <th className="p-1.5 text-center w-28">Gia công? (Cột 2)</th>
                            <th className="p-1.5 text-right w-44">Phí gia công (Cột 3)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {generatedForm.phan1.cau6_doanhThuSanPham.map((sp, sIdx) => (
                            <tr key={sIdx}>
                              <td className="p-1.5">{sp.tenNganh}</td>
                              <td className="p-1.5 font-mono">{sp.maIO}</td>
                              <td className="p-1.5 text-right font-mono font-bold">
                                {sp.c1_dtt.toLocaleString("vi-VN")}
                              </td>
                              <td className="p-1.5 text-center">
                                <span className={sp.c2_coGiaCong === 1 ? "text-emerald-700 font-bold" : "text-slate-400"}>
                                  {sp.c2_coGiaCong === 1 ? "Có" : "Không"}
                                </span>
                              </td>
                              <td className="p-1.5 text-right font-mono">
                                {sp.c3_phiGiaCong.toLocaleString("vi-VN")}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {/* Bảng Câu 8 (Tồn kho) */}
                    <div className="border border-slate-200 mt-2">
                      <div className="bg-slate-100 p-1.5 font-bold text-slate-800 border-b border-slate-200">
                        8. Hàng tồn kho và chi phí dở dang của doanh nghiệp năm 2025
                      </div>
                      <div className="grid grid-cols-3 gap-2 p-2 bg-slate-50/40 text-xs">
                        <div className="p-1.5 bg-white border border-slate-200">
                          <span className="text-[11px] text-slate-500 block">8.1 Chi phí SX dở dang (TK 154)</span>
                          <span className="font-mono font-bold text-slate-800">
                            Đầu kỳ: {generatedForm.phan1.cau8_hangTonKhoDN.c81_sxdd_dauKy.toLocaleString("vi-VN")} | Cuối kỳ:{" "}
                            {generatedForm.phan1.cau8_hangTonKhoDN.c81_sxdd_cuoiKy.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="p-1.5 bg-white border border-slate-200">
                          <span className="text-[11px] text-slate-500 block">8.2 Thành phẩm tồn kho (TK 155)</span>
                          <span className="font-mono font-bold text-slate-800">
                            Đầu kỳ: {generatedForm.phan1.cau8_hangTonKhoDN.c82_thanhPham_dauKy.toLocaleString("vi-VN")} | Cuối kỳ:{" "}
                            {generatedForm.phan1.cau8_hangTonKhoDN.c82_thanhPham_cuoiKy.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="p-1.5 bg-white border border-slate-200">
                          <span className="text-[11px] text-slate-500 block">8.3 Hàng gửi đi bán (TK 157)</span>
                          <span className="font-mono font-bold text-slate-800">
                            Đầu kỳ: {generatedForm.phan1.cau8_hangTonKhoDN.c83_hangGui_dauKy.toLocaleString("vi-VN")} | Cuối kỳ:{" "}
                            {generatedForm.phan1.cau8_hangTonKhoDN.c83_hangGui_cuoiKy.toLocaleString("vi-VN")}
                          </span>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* PHẦN II: NGÀNH SẢN PHẨM CHÍNH (CÂU 10) */}
                  <div className="border border-slate-200 p-3 bg-white space-y-2">
                    <span className="font-bold text-sky-900 block border-b border-slate-200 pb-1 uppercase">
                      PHẦN II: KẾT QUẢ HOẠT ĐỘNG SXKD NGÀNH SẢN PHẨM CHÍNH (CÂU 10)
                    </span>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                      <div className="p-2 border border-slate-200 bg-slate-50">
                        <span className="text-[11px] text-slate-500 block">Doanh thu thuần (DTT):</span>
                        <span className="font-mono font-bold text-slate-900">
                          {generatedForm.phan2.cau10_sxkdChinh.dtt.toLocaleString("vi-VN")}
                        </span>
                      </div>
                      <div className="p-2 border border-slate-200 bg-slate-50">
                        <span className="text-[11px] text-slate-500 block">Giá vốn hàng bán (GVHB):</span>
                        <span className="font-mono font-bold text-slate-900">
                          {generatedForm.phan2.cau10_sxkdChinh.gvhb.toLocaleString("vi-VN")}
                        </span>
                      </div>
                      <div className="p-2 border border-slate-200 bg-slate-50">
                        <span className="text-[11px] text-slate-500 block">Chi phí bán hàng (CFBH):</span>
                        <span className="font-mono font-bold text-slate-900">
                          {generatedForm.phan2.cau10_sxkdChinh.cfbh.toLocaleString("vi-VN")}
                        </span>
                      </div>
                      <div className="p-2 border border-slate-200 bg-slate-50">
                        <span className="text-[11px] text-slate-500 block">Chi phí quản lý (CFQL):</span>
                        <span className="font-mono font-bold text-slate-900">
                          {generatedForm.phan2.cau10_sxkdChinh.cfql.toLocaleString("vi-VN")}
                        </span>
                      </div>
                      <div className="p-2 border border-slate-200 bg-emerald-50/50">
                        <span className="text-[11px] text-emerald-800 font-bold block">Lợi nhuận SXKD (Mã 223):</span>
                        <span className="font-mono font-bold text-emerald-950">
                          {generatedForm.phan2.cau10_sxkdChinh.ma223_loiNhuan.toLocaleString("vi-VN")}
                        </span>
                      </div>
                      <div className="p-2 border border-slate-200 bg-slate-50">
                        <span className="text-[11px] text-slate-500 block">Trả lãi tiền vay (Mã 224):</span>
                        <span className="font-mono font-bold text-slate-900">
                          {generatedForm.phan2.cau10_sxkdChinh.ma224_laiVay.toLocaleString("vi-VN")}
                        </span>
                      </div>
                      <div className="p-2 border border-slate-200 bg-slate-50">
                        <span className="text-[11px] text-slate-500 block">Chi phí NVL trực tiếp:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {generatedForm.phan2.cau10_sxkdChinh.cfnvl.toLocaleString("vi-VN")}
                        </span>
                      </div>
                      <div className="p-2 border border-slate-200 bg-slate-50">
                        <span className="text-[11px] text-slate-500 block">Chi phí nhân công trực tiếp:</span>
                        <span className="font-mono font-bold text-slate-900">
                          {generatedForm.phan2.cau10_sxkdChinh.cfnc.toLocaleString("vi-VN")}
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* CÂU 15: BÓC TÁCH NGUYÊN VẬT LIỆU THEO TỪNG MÃ IO NGÀNH NGHỀ */}
                  <div className="border border-slate-200 p-3 bg-white space-y-2">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 border-b border-slate-200 pb-1">
                      <span className="font-bold text-sky-900 block uppercase">
                        15. BÓC TÁCH CHI PHÍ NGUYÊN VẬT LIỆU THEO MÃ IO (CÂU 15)
                      </span>
                      <span className="text-[11px] text-slate-500 font-medium">
                        Tổng giá trị NVL đã phân bổ:{" "}
                        <strong className="text-sky-900 font-mono">
                          {generatedForm.phan3.cau15_nvlChiTiet
                            .reduce((s, r) => s + r.c1_giaTri, 0)
                            .toLocaleString("vi-VN")}{" "}
                          đ
                        </strong>
                      </span>
                    </div>

                    <div className="overflow-x-auto border border-slate-200">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                            <th className="p-1.5 border-r border-slate-200 text-center w-16">Mã IO</th>
                            <th className="p-1.5 border-r border-slate-200">Tên nguyên liệu, vật liệu</th>
                            <th className="p-1.5 border-r border-slate-200 text-right w-44">Trị giá (đồng) (Cột 1)</th>
                            <th className="p-1.5 border-r border-slate-200 text-center w-28">Tỷ lệ NK (%) (Cột 2)</th>
                            <th className="p-1.5 text-right w-36">Gia công (Cột 3)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {generatedForm.phan3.cau15_nvlChiTiet.map((item, idx) => (
                            <tr key={`${item.maIO}-${idx}`} className="border-b border-slate-100 hover:bg-slate-50">
                              <td className="p-1.5 border-r border-slate-200 text-center font-mono font-bold text-sky-800 bg-sky-50/40">
                                {item.maIO}
                              </td>
                              <td className="p-1.5 border-r border-slate-200 text-slate-800 font-medium">
                                {item.moTa}
                              </td>
                              <td className="p-1.5 border-r border-slate-200 text-right font-mono">
                                <input
                                  type="number"
                                  value={item.c1_giaTri}
                                  onChange={e => handleUpdateCau15Item(idx, "c1_giaTri", Number(e.target.value))}
                                  className="w-full text-right bg-transparent border-0 border-b border-slate-200 hover:border-sky-500 focus:border-sky-600 focus:bg-white px-1 py-0.5 font-bold text-slate-900"
                                />
                              </td>
                              <td className="p-1.5 border-r border-slate-200 text-center font-mono">
                                <input
                                  type="number"
                                  min={0}
                                  max={100}
                                  value={item.c2_tyLeNK}
                                  onChange={e => handleUpdateCau15Item(idx, "c2_tyLeNK", Number(e.target.value))}
                                  className="w-16 text-center bg-transparent border-0 border-b border-slate-200 hover:border-sky-500 focus:border-sky-600 focus:bg-white px-1 py-0.5 text-slate-800"
                                />
                              </td>
                              <td className="p-1.5 text-right font-mono">
                                <input
                                  type="number"
                                  value={item.c3_giaCong}
                                  onChange={e => handleUpdateCau15Item(idx, "c3_giaCong", Number(e.target.value))}
                                  className="w-full text-right bg-transparent border-0 border-b border-slate-200 hover:border-sky-500 focus:border-sky-600 focus:bg-white px-1 py-0.5 text-slate-800"
                                />
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr className="bg-sky-50/70 font-bold border-t border-sky-200 text-sky-950">
                            <td colSpan={2} className="p-1.5 text-right border-r border-slate-200">
                              TỔNG CỘNG CHI PHÍ NGUYÊN VẬT LIỆU (CÂU 15):
                            </td>
                            <td className="p-1.5 text-right font-mono border-r border-slate-200 text-sky-900">
                              {generatedForm.phan3.cau15_nvlChiTiet
                                .reduce((s, r) => s + r.c1_giaTri, 0)
                                .toLocaleString("vi-VN")}
                            </td>
                            <td className="p-1.5 text-center text-slate-500 border-r border-slate-200">-</td>
                            <td className="p-1.5 text-right font-mono">
                              {generatedForm.phan3.cau15_nvlChiTiet
                                .reduce((s, r) => s + r.c3_giaCong, 0)
                                .toLocaleString("vi-VN")}
                            </td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>

                  {/* CÂU 18: BÓC TÁCH DỊCH VỤ MUA NGOÀI & CHI PHÍ KHÁC THEO TỪNG MÃ IO */}
                  <div className="border border-slate-200 p-3 bg-white space-y-2">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 border-b border-slate-200 pb-1">
                      <span className="font-bold text-sky-900 block uppercase">
                        18. BÓC TÁCH DỊCH VỤ MUA NGOÀI & CHI PHÍ KHÁC BẰNG TIỀN THEO MÃ IO (CÂU 18)
                      </span>
                      <span className="text-[11px] text-slate-500 font-medium">
                        Tổng dịch vụ mua ngoài:{" "}
                        <strong className="text-sky-900 font-mono">
                          {generatedForm.phan3.cau18_muaNgoaiChiPhiKhac
                            .reduce((s, r) => s + r.c1_giaTri, 0)
                            .toLocaleString("vi-VN")}{" "}
                          đ
                        </strong>
                      </span>
                    </div>

                    <div className="overflow-x-auto border border-slate-200">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                            <th className="p-1.5 border-r border-slate-200 text-center w-16">Mã IO</th>
                            <th className="p-1.5 border-r border-slate-200">Tên loại chi phí / dịch vụ mua ngoài</th>
                            <th className="p-1.5 border-r border-slate-200 text-right w-44">Trị giá (đồng) (Cột 1)</th>
                            <th className="p-1.5 text-center w-28">Tỷ lệ NK (%) (Cột 2)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {generatedForm.phan3.cau18_muaNgoaiChiPhiKhac.map((item, idx) => (
                            <tr key={`${item.maIO}-${idx}`} className="border-b border-slate-100 hover:bg-slate-50">
                              <td className="p-1.5 border-r border-slate-200 text-center font-mono font-bold text-emerald-800 bg-emerald-50/40">
                                {item.maIO}
                              </td>
                              <td className="p-1.5 border-r border-slate-200 text-slate-800 font-medium">
                                {item.moTa}
                              </td>
                              <td className="p-1.5 border-r border-slate-200 text-right font-mono">
                                <input
                                  type="number"
                                  value={item.c1_giaTri}
                                  onChange={e => handleUpdateCau18Item(idx, "c1_giaTri", Number(e.target.value))}
                                  className="w-full text-right bg-transparent border-0 border-b border-slate-200 hover:border-emerald-500 focus:border-emerald-600 focus:bg-white px-1 py-0.5 font-bold text-slate-900"
                                />
                              </td>
                              <td className="p-1.5 text-center font-mono">
                                <input
                                  type="number"
                                  min={0}
                                  max={100}
                                  value={item.c2_tyLeNK}
                                  onChange={e => handleUpdateCau18Item(idx, "c2_tyLeNK", Number(e.target.value))}
                                  className="w-16 text-center bg-transparent border-0 border-b border-slate-200 hover:border-emerald-500 focus:border-emerald-600 focus:bg-white px-1 py-0.5 text-slate-800"
                                />
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr className="bg-emerald-50/70 font-bold border-t border-emerald-200 text-emerald-950">
                            <td colSpan={2} className="p-1.5 text-right border-r border-slate-200">
                              TỔNG CỘNG CHI PHÍ DỊCH VỤ MUA NGOÀI (CÂU 18):
                            </td>
                            <td className="p-1.5 text-right font-mono text-emerald-900">
                              {generatedForm.phan3.cau18_muaNgoaiChiPhiKhac
                                .reduce((s, r) => s + r.c1_giaTri, 0)
                                .toLocaleString("vi-VN")}
                            </td>
                            <td className="p-1.5 text-center text-slate-500">-</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>

                  {/* PHẦN III & IV: CHI PHÍ NHÂN CÔNG & KHẤU HAO & VAY */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div className="border border-slate-200 p-3 bg-white space-y-2">
                      <span className="font-bold text-sky-900 block border-b border-slate-200 pb-1 uppercase">
                        PHẦN III: CHI PHÍ NHÂN CÔNG (CÂU 16) & KHẤU HAO (CÂU 17)
                      </span>
                      <div className="space-y-1 text-xs">
                        <div className="flex justify-between py-1 border-b border-slate-100">
                          <span className="text-slate-600">Tiền lương, phụ cấp (Mã 182 - TK 334):</span>
                          <span className="font-mono font-bold">
                            {generatedForm.phan3.cau16_nhanCong.ma182_tienLuong.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="flex justify-between py-1 border-b border-slate-100">
                          <span className="text-slate-600">Bảo hiểm xã hội (Mã 183):</span>
                          <span className="font-mono font-bold">
                            {generatedForm.phan3.cau16_nhanCong.ma183_bhxh.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="flex justify-between py-1 border-b border-slate-100">
                          <span className="text-slate-600">Bảo hiểm y tế (Mã 184):</span>
                          <span className="font-mono font-bold">
                            {generatedForm.phan3.cau16_nhanCong.ma184_bhyt.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="flex justify-between py-1 border-b border-slate-100">
                          <span className="text-slate-600">Kinh phí công đoàn (Mã 187):</span>
                          <span className="font-mono font-bold">
                            {generatedForm.phan3.cau16_nhanCong.ma187_kpcd.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="flex justify-between py-1 bg-slate-50 font-bold px-1">
                          <span>Chi phí khấu hao TSCĐ (Mã 225 - TK 214):</span>
                          <span className="font-mono text-indigo-900">
                            {generatedForm.phan3.cau17_khauHao225.toLocaleString("vi-VN")}
                          </span>
                        </div>
                      </div>
                    </div>

                    <div className="border border-slate-200 p-3 bg-white space-y-2">
                      <span className="font-bold text-sky-900 block border-b border-slate-200 pb-1 uppercase">
                        PHẦN IV & TỔNG HỢP HỆ SỐ IO (CÂU 19 & TỶ LỆ IC/GO)
                      </span>
                      <div className="space-y-1 text-xs">
                        <div className="flex justify-between py-1 border-b border-slate-100">
                          <span className="text-slate-600">Vay ngắn hạn (Mã 320 CĐKT):</span>
                          <span className="font-mono font-bold">
                            {generatedForm.phan4.cau194_vayNganHan320.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="flex justify-between py-1 border-b border-slate-100">
                          <span className="text-slate-600">Vay dài hạn (Mã 338 CĐKT):</span>
                          <span className="font-mono font-bold">
                            {generatedForm.phan4.cau194_vayDaiHan338.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="flex justify-between py-1 border-b border-slate-100">
                          <span className="text-slate-600">Chi phí trung gian (IC):</span>
                          <span className="font-mono font-bold text-emerald-800">
                            {generatedForm.tongHopIO.ic.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="flex justify-between py-1 border-b border-slate-100">
                          <span className="text-slate-600">Giá trị sản xuất (GO):</span>
                          <span className="font-mono font-bold text-emerald-800">
                            {generatedForm.tongHopIO.go.toLocaleString("vi-VN")}
                          </span>
                        </div>
                        <div className="flex justify-between py-1 bg-emerald-50 px-1 font-bold">
                          <span className="text-emerald-950">Tỷ lệ IC/GO sơ bộ:</span>
                          <span className="font-mono text-emerald-900 text-sm">
                            {(generatedForm.tongHopIO.tyLeIC_GO * 100).toFixed(1)}% ({generatedForm.tongHopIO.danhGiaIC})
                          </span>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ======================================================== */}
      {/* CHỨC NĂNG 2: KIỂM TRA LỖI DANH SÁCH DOANH NGHIỆP HÀNG LOẠT */}
      {/* ======================================================== */}
      {activeSubTab === "batch_audit" && (
        <div className="px-3 sm:px-6 space-y-4">
          {/* THANH THỐNG KÊ TỔNG HỢP VÀ TẢI FILE */}
          <div className="bg-slate-50 border border-slate-300 p-3.5 space-y-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200 pb-2">
              <div>
                <h3 className="text-sm font-bold text-slate-900 m-0 flex items-center gap-1.5">
                  <ShieldAlert className="w-4 h-4 text-rose-600" />
                  QUÉT & TỔNG HỢP LỖI CỦA DANH SÁCH DOANH NGHIỆP THEO HƯỚNG DẪN IO-2026
                </h3>
                <p className="text-xs text-slate-500 m-0">
                  Đang làm việc: <strong>{batchData.length} doanh nghiệp</strong> ({batchFileName})
                </p>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => batchExcelInputRef.current?.click()}
                  className="bg-sky-600 hover:bg-sky-700 text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer border-0"
                >
                  <Upload className="w-3.5 h-3.5" /> Nạp danh sách DN từ Excel
                </button>
                <button
                  type="button"
                  onClick={() => exportBatchAuditExcel(batchAuditResults)}
                  disabled={batchAuditResults.length === 0}
                  className="bg-emerald-700 hover:bg-emerald-800 disabled:opacity-50 text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer border-0"
                >
                  <Download className="w-3.5 h-3.5" /> Xuất Báo Cáo Giải Trình (Excel)
                </button>
                <input
                  type="file"
                  ref={batchExcelInputRef}
                  onChange={handleBatchFileUpload}
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                />
              </div>
            </div>

            {/* THẺ CHỈ SỐ KPI */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
              <div className="bg-white border border-slate-300 p-2.5">
                <span className="text-[11px] text-slate-500 font-bold block uppercase">Tổng số DN quét</span>
                <span className="text-lg font-bold text-slate-900">{batchStats.total}</span>
              </div>
              <div className="bg-white border border-rose-300 p-2.5">
                <span className="text-[11px] text-rose-700 font-bold block uppercase">DN có Lỗi (Bắt buộc sửa)</span>
                <span className="text-lg font-bold text-rose-700">{batchStats.errorCount}</span>
              </div>
              <div className="bg-white border border-amber-300 p-2.5">
                <span className="text-[11px] text-amber-700 font-bold block uppercase">DN có Cảnh báo (Cần ĐTV)</span>
                <span className="text-lg font-bold text-amber-700">{batchStats.warningCount}</span>
              </div>
              <div className="bg-white border border-emerald-300 p-2.5">
                <span className="text-[11px] text-emerald-700 font-bold block uppercase">DN Đạt chuẩn hợp lệ</span>
                <span className="text-lg font-bold text-emerald-700">{batchStats.validCount}</span>
              </div>
            </div>

            {/* BỘ LỌC TÌM KIẾM */}
            <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
              <div className="flex items-center gap-1 text-xs">
                <span className="font-bold text-slate-600">Lọc theo:</span>
                <button
                  type="button"
                  onClick={() => setFilterMode("all")}
                  className={`px-2.5 py-1 text-xs font-bold border cursor-pointer ${
                    filterMode === "all" ? "bg-slate-800 text-white border-slate-800" : "bg-white text-slate-700 border-slate-300"
                  }`}
                >
                  Tất cả ({batchStats.total})
                </button>
                <button
                  type="button"
                  onClick={() => setFilterMode("errors_only")}
                  className={`px-2.5 py-1 text-xs font-bold border cursor-pointer ${
                    filterMode === "errors_only"
                      ? "bg-rose-600 text-white border-rose-600"
                      : "bg-white text-rose-700 border-rose-300"
                  }`}
                >
                  Chỉ DN có Lỗi ({batchStats.errorCount})
                </button>
                <button
                  type="button"
                  onClick={() => setFilterMode("warnings_only")}
                  className={`px-2.5 py-1 text-xs font-bold border cursor-pointer ${
                    filterMode === "warnings_only"
                      ? "bg-amber-500 text-slate-950 border-amber-500"
                      : "bg-white text-amber-800 border-amber-300"
                  }`}
                >
                  Chỉ DN có Cảnh báo ({batchStats.warningCount})
                </button>
              </div>

              <div className="relative">
                <input
                  type="text"
                  value={searchDn}
                  onChange={e => setSearchDn(e.target.value)}
                  placeholder="🔍 Tìm theo MST hoặc Tên DN..."
                  className="bg-white border border-slate-300 text-xs px-2.5 py-1 w-64 focus:outline-none focus:border-sky-500"
                />
                {searchDn && (
                  <button
                    type="button"
                    onClick={() => setSearchDn("")}
                    className="absolute right-2 top-1 text-slate-400 hover:text-slate-800 text-xs border-0 bg-transparent cursor-pointer"
                  >
                    ✕
                  </button>
                )}
              </div>
            </div>
          </div>

          {/* BẢNG KẾT QUẢ QUÉT HÀNG LOẠT */}
          <div className="border border-slate-300 overflow-x-auto bg-white">
            <table className="w-full text-xs text-left border-collapse">
              <thead className="bg-[#286e42] text-white font-bold sticky top-0">
                <tr>
                  <th className="py-2 px-2.5 text-center w-12 border-r border-emerald-700">STT</th>
                  <th className="py-2 px-3 w-32 border-r border-emerald-700">Mã số thuế</th>
                  <th className="py-2 px-3 border-r border-emerald-700">Tên doanh nghiệp</th>
                  <th className="py-2 px-2.5 text-center w-28 border-r border-emerald-700">Trạng thái</th>
                  <th className="py-2 px-2 text-center w-16 border-r border-emerald-700">Số lỗi</th>
                  <th className="py-2 px-2 text-center w-20 border-r border-emerald-700">Cảnh báo</th>
                  <th className="py-2 px-3 text-right w-36 border-r border-emerald-700">Doanh thu (C5)</th>
                  <th className="py-2 px-2 text-center w-24 border-r border-emerald-700">Tỷ lệ IC/GO</th>
                  <th className="py-2 px-3">Chi tiết lỗi vi phạm (Đề nghị ĐTV giải trình)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredBatchResults.length === 0 ? (
                  <tr>
                    <td colSpan={9} className="py-8 text-center text-slate-400 italic">
                      Không tìm thấy doanh nghiệp nào theo điều kiện lọc.
                    </td>
                  </tr>
                ) : (
                  filteredBatchResults.map((item, idx) => (
                    <tr
                      key={item.mst || idx}
                      className={
                        item.soLoi > 0
                          ? "bg-rose-50/40 hover:bg-rose-50"
                          : item.soCanhBao > 0
                          ? "bg-amber-50/30 hover:bg-amber-50/60"
                          : idx % 2 === 1
                          ? "bg-slate-50/40 hover:bg-slate-50"
                          : "hover:bg-slate-50"
                      }
                    >
                      <td className="py-1.5 px-2.5 text-center font-bold text-slate-500 font-mono border-r border-slate-100">
                        {idx + 1}
                      </td>
                      <td className="py-1.5 px-3 font-mono font-bold text-slate-900 border-r border-slate-100">
                        {item.mst}
                      </td>
                      <td className="py-1.5 px-3 font-semibold text-slate-800 border-r border-slate-100">
                        {item.tenDN}
                      </td>
                      <td className="py-1.5 px-2.5 text-center border-r border-slate-100">
                        {item.soLoi > 0 ? (
                          <span className="inline-block px-1.5 py-0.5 text-[10px] font-bold bg-rose-100 text-rose-800 border border-rose-300">
                            CÓ LỖI
                          </span>
                        ) : item.soCanhBao > 0 ? (
                          <span className="inline-block px-1.5 py-0.5 text-[10px] font-bold bg-amber-100 text-amber-900 border border-amber-300">
                            CẢNH BÁO
                          </span>
                        ) : (
                          <span className="inline-block px-1.5 py-0.5 text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">
                            HỢP LỆ
                          </span>
                        )}
                      </td>
                      <td className="py-1.5 px-2 text-center font-mono font-bold text-rose-700 border-r border-slate-100">
                        {item.soLoi}
                      </td>
                      <td className="py-1.5 px-2 text-center font-mono font-bold text-amber-700 border-r border-slate-100">
                        {item.soCanhBao}
                      </td>
                      <td className="py-1.5 px-3 text-right font-mono text-slate-800 border-r border-slate-100">
                        {item.dtt.toLocaleString("vi-VN")}
                      </td>
                      <td className="py-1.5 px-2 text-center font-mono font-semibold text-slate-700 border-r border-slate-100">
                        {(item.tyLeIC_GO * 100).toFixed(1)}%
                      </td>
                      <td className="py-1.5 px-3">
                        {item.violations.length === 0 ? (
                          <span className="text-emerald-700 font-medium text-[11px]">Đạt chuẩn, không có lỗi</span>
                        ) : (
                          <div className="space-y-0.5">
                            {item.violations.map((v, vIdx) => (
                              <div key={vIdx} className="text-[11px]">
                                <span
                                  className={
                                    v.type === "error"
                                      ? "text-rose-700 font-bold"
                                      : "text-amber-800 font-semibold"
                                  }
                                >
                                  [{v.cau}] {v.tenLoi}
                                </span>
                              </div>
                            ))}
                          </div>
                        )}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};

export default EnterpriseAuditAndBctcBridge;
