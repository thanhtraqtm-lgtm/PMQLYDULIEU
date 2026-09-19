// Bộ giải mã và bóc tách Báo cáo Tài chính XML từ Tổng cục Thuế / HTKK (TT133, TT200)
import * as XLSX from "xlsx";

export interface BctcHeaderInfo {
  mst: string;
  tenDoanhNghiep: string;
  diaChi: string;
  tinhTP: string;
  maTKhai: string;
  tenTKhai: string;
  moTaBMau: string;
  pbanTKhaiXML: string;
  loaiTKhai: string;
  soLan: number;
  kyKKhai: string;
  kyKKhaiTuNgay: string;
  kyKKhaiDenNgay: string;
  maCQT: string;
  tenCQT: string;
  ngayLapTKhai: string;
  ngayKy: string;
  bctcDaKiemToan: boolean;
  ngayLap: string;
}

export interface BctcBalanceItem {
  maSo: string;
  tenChiTieu: string;
  thuyetMinh: string;
  soCuoiNam: number;
  soDauNam: number;
  chenhLech: number;
  phanTramThayDoi: number;
  isHeader?: boolean;
}

export interface BctcProfitLossItem {
  maSo: string;
  tenChiTieu: string;
  thuyetMinh: string;
  namNay: number;
  namTruoc: number;
  chenhLech: number;
  phanTramThayDoi: number;
}

export interface BctcCashFlowItem {
  maSo: string;
  tenChiTieu: string;
  thuyetMinh: string;
  namNay: number;
  namTruoc: number;
  chenhLech: number;
}

export interface BctcTrialBalanceItem {
  soHieuTK: string;
  tenTK: string;
  duNoDau: number;
  duCoDau: number;
  psNo: number;
  psCo: number;
  duNoCuoi: number;
  duCoCuoi: number;
  capTK: number;
}

export interface BctcFinancialRatios {
  thanhToanHienHanh: number; // TSNH / NoNH
  thanhToanNhanh: number; // (Tien + PhaiThu) / NoNH
  tyLeTienMat: number; // Tien / NoNH
  heSoNoTrenTongTaiSan: number; // No / TTS
  heSoNoTrenVCSH: number; // No / VCSH
  heSoTuTaiTro: number; // VCSH / TTS
  bienLoiNhuanGop: number; // LN Gop / DTT (%)
  bienLoiNhuanRong: number; // LNST / DTT (%)
  roa: number; // LNST / Tong Tai San (%)
  roe: number; // LNST / VCSH (%)
  vongQuayTongTaiSan: number; // DTT / TTS
  vongQuayHangTonKho: number; // GVHB / TonKho
}

export interface BctcParsedData {
  header: BctcHeaderInfo;
  balanceSheet: BctcBalanceItem[];
  profitLoss: BctcProfitLossItem[];
  cashFlow: BctcCashFlowItem[];
  trialBalance: BctcTrialBalanceItem[];
  ratios: BctcFinancialRatios;
  rawSummary: {
    tongTaiSanCuoiNam: number;
    tongTaiSanDauNam: number;
    noPhaiTraCuoiNam: number;
    vonChuSoHuuCuoiNam: number;
    doanhThuThuan: number;
    giaVon: number;
    loiNhuanGop: number;
    loiNhuanTruocThue: number;
    loiNhuanSauThue: number;
    tienVaTuongDuongTienCuoiKy: number;
    tongDuNoCuoiKy: number;
    tongDuCoCuoiKy: number;
    isTrialBalanceBalanced: boolean;
  };
}

// Từ điển tên chỉ tiêu Bảng cân đối kế toán TT133 (Mẫu B01a - DNN)
const DICT_CDKT_TT133: Record<string, string> = {
  "110": "I. Tiền và các khoản tương đương tiền",
  "120": "II. Đầu tư tài chính",
  "121": "1. Chứng khoán kinh doanh",
  "122": "2. Đầu tư nắm giữ đến ngày đáo hạn",
  "123": "3. Dự phòng giảm giá đầu tư tài chính (*)",
  "130": "III. Các khoản phải thu",
  "131": "1. Phải thu của khách hàng",
  "132": "2. Trả trước cho người bán",
  "133": "3. Vốn kinh doanh ở đơn vị trực thuộc",
  "134": "4. Phải thu khác",
  "135": "5. Dự phòng phải thu khó đòi (*)",
  "136": "6. Tài sản thiếu chờ xử lý",
  "140": "IV. Hàng tồn kho",
  "141": "1. Hàng tồn kho",
  "142": "2. Dự phòng giảm giá hàng tồn kho (*)",
  "150": "V. Tài sản cố định",
  "151": "1. Nguyên giá TSCĐ",
  "152": "2. Giá trị hao mòn lũy kế (*)",
  "160": "VI. Bất động sản đầu tư",
  "161": "1. Nguyên giá BĐS đầu tư",
  "162": "2. Giá trị hao mòn lũy kế (*)",
  "170": "VII. Xây dựng cơ bản dở dang",
  "180": "VIII. Tài sản khác",
  "181": "1. Thuế GTGT được khấu trừ",
  "182": "2. Tài sản khác",
  "200": "TỔNG CỘNG TÀI SẢN (200 = 110+120+130+140+150+160+170+180)",
  "300": "A. NỢ PHẢI TRẢ (300 = 311+312+...+320)",
  "311": "1. Phải trả người bán",
  "312": "2. Người mua trả tiền trước",
  "313": "3. Thuế và các khoản phải nộp Nhà nước",
  "314": "4. Phải trả người lao động",
  "315": "5. Chi phí phải trả",
  "316": "6. Phải trả nội bộ",
  "317": "7. Doanh thu chưa thực hiện",
  "318": "8. Phải trả khác",
  "319": "9. Vay và nợ thuê tài chính",
  "320": "10. Dự phòng phải trả",
  "400": "B. VỐN CHỦ SỞ HỮU (400 = 411+...+417)",
  "411": "1. Vốn đầu tư của chủ sở hữu",
  "412": "2. Thặng dư vốn cổ phần",
  "413": "3. Vốn khác của chủ sở hữu",
  "414": "4. Cổ phiếu quỹ (*)",
  "415": "5. Chênh lệch tỷ giá hối đoái",
  "416": "6. Các quỹ thuộc vốn chủ sở hữu",
  "417": "7. Lợi nhuận sau thuế chưa phân phối",
  "500": "TỔNG CỘNG NGUỒN VỐN (500 = 300 + 400)"
};

// Từ điển tên chỉ tiêu Báo cáo Kết quả kinh doanh (B02 - DNN TT133)
const DICT_KQKD_TT133: Record<string, string> = {
  "01": "1. Doanh thu bán hàng và cung cấp dịch vụ",
  "02": "2. Các khoản giảm trừ doanh thu",
  "10": "3. Doanh thu thuần về bán hàng và cung cấp dịch vụ (10 = 01 - 02)",
  "11": "4. Giá vốn hàng bán",
  "20": "5. Lợi nhuận gộp về bán hàng và cung cấp dịch vụ (20 = 10 - 11)",
  "21": "6. Doanh thu hoạt động tài chính",
  "22": "7. Chi phí tài chính",
  "23": "- Trong đó: Chi phí lãi vay",
  "24": "8. Chi phí quản lý kinh doanh",
  "30": "9. Lợi nhuận thuần từ hoạt động kinh doanh (30 = 20 + 21 - 22 - 24)",
  "31": "10. Thu nhập khác",
  "32": "11. Chi phí khác",
  "40": "12. Lợi nhuận khác (40 = 31 - 32)",
  "50": "13. Tổng lợi nhuận kế toán trước thuế (50 = 30 + 40)",
  "51": "14. Chi phí thuế TNDN",
  "60": "15. Lợi nhuận sau thuế thu nhập doanh nghiệp (60 = 50 - 51)"
};

// Từ điển tên chỉ tiêu Báo cáo Lưu chuyển tiền tệ trực tiếp (TT133)
const DICT_LCTT_TT133: Record<string, string> = {
  "01": "1. Tiền thu từ bán hàng, cung cấp DV và DT khác",
  "02": "2. Tiền chi trả cho người cung cấp hàng hóa và DV",
  "03": "3. Tiền chi trả cho người lao động",
  "04": "4. Tiền lãi vay đã trả",
  "05": "5. Thuế TNDN đã nộp",
  "06": "6. Tiền thu khác từ hoạt động kinh doanh",
  "07": "7. Tiền chi khác cho hoạt động kinh doanh",
  "20": "I. Lưu chuyển tiền thuần từ hoạt động kinh doanh",
  "21": "1. Tiền chi để mua sắm, xây dựng TSCĐ và TSDH khác",
  "22": "2. Tiền thu từ thanh lý, nhượng bán TSCĐ và TSDH khác",
  "23": "3. Tiền chi cho vay, mua các công cụ nợ của đơn vị khác",
  "24": "4. Tiền thu hồi cho vay, bán lại công cụ nợ của đơn vị khác",
  "25": "5. Tiền chi đầu tư góp vốn vào đơn vị khác",
  "26": "6. Tiền thu hồi đầu tư góp vốn vào đơn vị khác",
  "27": "7. Tiền thu lãi cho vay, cổ tức và lợi nhuận được chia",
  "30": "II. Lưu chuyển tiền thuần từ hoạt động đầu tư",
  "31": "1. Tiền thu từ phát hành cổ phiếu, nhận vốn góp của CSH",
  "32": "2. Tiền trả lại vốn góp cho các CSH, mua lại CP đã phát hành",
  "33": "3. Tiền thu từ đi vay",
  "34": "4. Tiền trả nợ gốc vay",
  "35": "5. Tiền trả nợ gốc thuê tài chính",
  "36": "6. Cổ tức, lợi nhuận đã trả cho chủ sở hữu",
  "40": "III. Lưu chuyển tiền thuần từ hoạt động tài chính",
  "50": "LƯU CHUYỂN TIỀN THUẦN TRONG KỲ (50 = 20 + 30 + 40)",
  "60": "Tiền và tương đương tiền đầu kỳ",
  "61": "Ảnh hưởng của thay đổi tỷ giá hối đoái",
  "70": "Tiền và tương đương tiền cuối kỳ (70 = 50 + 60 + 61)"
};

// Từ điển danh mục tài khoản kế toán Việt Nam (TT133 & TT200)
const DICT_TAI_KHOAN: Record<string, string> = {
  "111": "Tiền mặt",
  "1111": "Tiền Việt Nam",
  "1112": "Ngoại tệ",
  "112": "Tiền gửi ngân hàng",
  "1121": "Tiền gửi ngân hàng (VND)",
  "1122": "Tiền gửi ngân hàng (Ngoại tệ)",
  "121": "Chứng khoán kinh doanh",
  "128": "Đầu tư nắm giữ đến ngày đáo hạn",
  "1281": "Tiền gửi có kỳ hạn",
  "1288": "Đầu tư nắm giữ đến ngày đáo hạn khác",
  "131": "Phải thu của khách hàng",
  "133": "Thuế GTGT được khấu trừ",
  "1331": "Thuế GTGT được khấu trừ của HHTT, DV",
  "1332": "Thuế GTGT được khấu trừ của TSCĐ",
  "136": "Phải thu nội bộ",
  "1361": "Vốn kinh doanh ở các đơn vị trực thuộc",
  "1368": "Phải thu nội bộ khác",
  "138": "Phải thu khác",
  "1381": "Tài sản thiếu chờ xử lý",
  "1386": "Cầm cố, thế chấp, ký quỹ, ký cược",
  "1388": "Phải thu khác",
  "141": "Tạm ứng",
  "151": "Hàng mua đang đi đường",
  "152": "Nguyên liệu, vật liệu",
  "153": "Công cụ, dụng cụ",
  "154": "Chi phí sản xuất, kinh doanh dở dang",
  "155": "Thành phẩm",
  "156": "Hàng hóa",
  "157": "Hàng gửi đi bán",
  "211": "Tài sản cố định hữu hình",
  "2111": "Nhà cửa, vật kiến trúc",
  "2112": "Máy móc, thiết bị",
  "2113": "Phương tiện vận tải, truyền dẫn",
  "214": "Hao mòn tài sản cố định",
  "2141": "Hao mòn TSCĐ hữu hình",
  "2142": "Hao mòn TSCĐ vô hình",
  "2143": "Hao mòn BĐS đầu tư",
  "2147": "Hao mòn TSCĐ khác",
  "217": "Bất động sản đầu tư",
  "228": "Đầu tư góp vốn vào đơn vị khác",
  "2281": "Đầu tư vào công ty con",
  "2288": "Đầu tư khác",
  "229": "Dự phòng tổn thất tài sản",
  "2291": "Dự phòng giảm giá chứng khoán kinh doanh",
  "2292": "Dự phòng tổn thất đầu tư vào đơn vị khác",
  "2293": "Dự phòng phải thu khó đòi",
  "2294": "Dự phòng giảm giá hàng tồn kho",
  "241": "Xây dựng cơ bản dở dang",
  "2411": "Mua sắm TSCĐ",
  "2412": "Xây dựng cơ bản",
  "2413": "Sửa chữa lớn TSCĐ",
  "242": "Chi phí trả trước",
  "331": "Phải trả cho người bán",
  "333": "Thuế và các khoản phải nộp Nhà nước",
  "3331": "Thuế giá trị gia tăng phải nộp",
  "33311": "Thuế GTGT đầu ra",
  "33312": "Thuế GTGT hàng nhập khẩu",
  "3332": "Thuế tiêu thụ đặc biệt",
  "3333": "Thuế xuất, nhập khẩu",
  "3334": "Thuế thu nhập doanh nghiệp",
  "3335": "Thuế thu nhập cá nhân",
  "3336": "Thuế tài nguyên",
  "3337": "Thuế nhà đất, tiền thuê đất",
  "3338": "Các loại thuế khác",
  "33381": "Thuế bảo vệ môi trường",
  "33382": "Các loại thuế khác",
  "3339": "Phí, lệ phí và các khoản phải nộp khác",
  "334": "Phải trả người lao động",
  "335": "Chi phí phải trả",
  "336": "Phải trả nội bộ",
  "3361": "Phải trả nội bộ về vốn kinh doanh",
  "3368": "Phải trả nội bộ khác",
  "338": "Phải trả, phải nộp khác",
  "3381": "Tài sản thừa chờ giải quyết",
  "3382": "Kinh phí công đoàn",
  "3383": "Bảo hiểm xã hội",
  "3384": "Bảo hiểm y tế",
  "3385": "Bảo hiểm thất nghiệp",
  "3386": "Nhận ký quỹ, ký cược",
  "3387": "Doanh thu chưa thực hiện",
  "3388": "Phải trả, phải nộp khác",
  "341": "Vay và nợ thuê tài chính",
  "3411": "Các khoản đi vay",
  "3412": "Nợ thuê tài chính",
  "352": "Dự phòng phải trả",
  "3521": "Dự phòng bảo hành sản phẩm, hàng hóa",
  "3522": "Dự phòng bảo hành công trình xây dựng",
  "3524": "Dự phòng phải trả khác",
  "353": "Quỹ khen thưởng, phúc lợi",
  "3531": "Quỹ khen thưởng",
  "3532": "Quỹ phúc lợi",
  "3533": "Quỹ phúc lợi đã hình thành TSCĐ",
  "3534": "Quỹ thưởng ban quản lý điều hành",
  "356": "Quỹ phát triển khoa học và công nghệ",
  "3561": "Quỹ phát triển KH và CN",
  "3562": "Quỹ phát triển KH và CN đã hình thành TSCĐ",
  "411": "Vốn đầu tư của chủ sở hữu",
  "4111": "Vốn góp của chủ sở hữu",
  "4112": "Thặng dư vốn cổ phần",
  "4118": "Vốn khác",
  "413": "Chênh lệch tỷ giá hối đoái",
  "418": "Các quỹ thuộc vốn chủ sở hữu",
  "419": "Cổ phiếu quỹ",
  "421": "Lợi nhuận sau thuế chưa phân phối",
  "4211": "Lợi nhuận sau thuế chưa phân phối năm trước",
  "4212": "Lợi nhuận sau thuế chưa phân phối năm nay",
  "511": "Doanh thu bán hàng và cung cấp dịch vụ",
  "5111": "Doanh thu bán hàng hóa",
  "5112": "Doanh thu bán các thành phẩm",
  "5113": "Doanh thu cung cấp dịch vụ",
  "5118": "Doanh thu khác",
  "515": "Doanh thu hoạt động tài chính",
  "611": "Mua hàng",
  "631": "Giá thành sản xuất",
  "632": "Giá vốn hàng bán",
  "635": "Chi phí tài chính",
  "642": "Chi phí quản lý kinh doanh",
  "6421": "Chi phí bán hàng",
  "6422": "Chi phí quản lý doanh nghiệp",
  "711": "Thu nhập khác",
  "811": "Chi phí khác",
  "821": "Chi phí thuế thu nhập doanh nghiệp",
  "911": "Xác định kết quả kinh doanh"
};

// Hàm trợ giúp đọc nội dung thẻ an toàn, bỏ qua tiền tố hoặc namespace
function getNodeText(parent: Element | Document, tagName: string): string {
  // 1. Thử trực tiếp
  const direct = parent.getElementsByTagName(tagName)[0];
  if (direct && direct.textContent !== null) return direct.textContent.trim();

  // 2. Thử lowercase
  const lowerTag = tagName.toLowerCase();
  const allElements = parent.getElementsByTagName("*");
  for (let i = 0; i < allElements.length; i++) {
    const el = allElements[i];
    const localName = el.localName || el.nodeName.split(":").pop() || "";
    if (localName.toLowerCase() === lowerTag && el.textContent !== null) {
      return el.textContent.trim();
    }
  }
  return "";
}

function getNodeNumber(parent: Element | Document, tagName: string): number {
  const text = getNodeText(parent, tagName);
  if (!text) return 0;
  const cleaned = text.replace(/,/g, "").replace(/\s/g, "");
  const val = Number(cleaned);
  return isNaN(val) ? 0 : val;
}

// Hàm chính bóc tách XML BCTC từ chuỗi văn bản
export function parseBctcXml(xmlString: string): BctcParsedData {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlString, "application/xml");

  const parseError = xmlDoc.getElementsByTagName("parsererror")[0];
  if (parseError) {
    throw new Error(`Lỗi cú pháp XML BCTC: ${parseError.textContent || "Không thể phân tích cú pháp XML"}`);
  }

  // 1. Thông tin chung tờ khai
  const header: BctcHeaderInfo = {
    mst: getNodeText(xmlDoc, "mst"),
    tenDoanhNghiep: getNodeText(xmlDoc, "tenNNT") || getNodeText(xmlDoc, "TenDoanhNghiep"),
    diaChi: getNodeText(xmlDoc, "dchiNNT") || getNodeText(xmlDoc, "DiaChi"),
    tinhTP: getNodeText(xmlDoc, "tenTinhNNT") || getNodeText(xmlDoc, "TinhTP"),
    maTKhai: getNodeText(xmlDoc, "maTKhai"),
    tenTKhai: getNodeText(xmlDoc, "tenTKhai"),
    moTaBMau: getNodeText(xmlDoc, "moTaBMau"),
    pbanTKhaiXML: getNodeText(xmlDoc, "pbanTKhaiXML"),
    loaiTKhai: getNodeText(xmlDoc, "loaiTKhai"),
    soLan: getNodeNumber(xmlDoc, "soLan"),
    kyKKhai: getNodeText(xmlDoc, "kyKKhai"),
    kyKKhaiTuNgay: getNodeText(xmlDoc, "kyKKhaiTuNgay"),
    kyKKhaiDenNgay: getNodeText(xmlDoc, "kyKKhaiDenNgay"),
    maCQT: getNodeText(xmlDoc, "maCQTNoiNop"),
    tenCQT: getNodeText(xmlDoc, "tenCQTNoiNop"),
    ngayLapTKhai: getNodeText(xmlDoc, "ngayLapTKhai"),
    ngayKy: getNodeText(xmlDoc, "ngayKy"),
    bctcDaKiemToan: getNodeText(xmlDoc, "bctcDaKiemToan") === "1",
    ngayLap: getNodeText(xmlDoc, "ngayLap")
  };

  // 2. Bóc tách Báo cáo tình hình tài chính / Cân đối kế toán (CTieuTKhaiChinh)
  const balanceSheet: BctcBalanceItem[] = [];
  const cTieuTKhaiChinh = xmlDoc.getElementsByTagName("CTieuTKhaiChinh")[0];

  if (cTieuTKhaiChinh) {
    const soCuoiNamEl = cTieuTKhaiChinh.getElementsByTagName("SoCuoiNam")[0];
    const soDauNamEl = cTieuTKhaiChinh.getElementsByTagName("SoDauNam")[0];
    const thuyetMinhEl = cTieuTKhaiChinh.getElementsByTagName("ThuyetMinh")[0];

    // Lấy tập hợp tất cả các mã chỉ tiêu xuất hiện trong SoCuoiNam hoặc SoDauNam
    const codeSet = new Set<string>();

    if (soCuoiNamEl) {
      for (let i = 0; i < soCuoiNamEl.children.length; i++) {
        const tag = soCuoiNamEl.children[i].localName || soCuoiNamEl.children[i].tagName;
        if (tag.toLowerCase().startsWith("ct")) {
          codeSet.add(tag.toLowerCase().substring(2));
        }
      }
    }
    if (soDauNamEl) {
      for (let i = 0; i < soDauNamEl.children.length; i++) {
        const tag = soDauNamEl.children[i].localName || soDauNamEl.children[i].tagName;
        if (tag.toLowerCase().startsWith("ct")) {
          codeSet.add(tag.toLowerCase().substring(2));
        }
      }
    }

    // Sắp xếp mã chỉ tiêu theo thứ tự logic bảng CĐKT
    const allCodes = Array.from(codeSet).sort((a, b) => {
      return parseInt(a, 10) - parseInt(b, 10);
    });

    for (const code of allCodes) {
      const tag = `ct${code}`;
      const cuoiNam = soCuoiNamEl ? getNodeNumber(soCuoiNamEl, tag) : 0;
      const dauNam = soDauNamEl ? getNodeNumber(soDauNamEl, tag) : 0;
      const tm = thuyetMinhEl ? getNodeText(thuyetMinhEl, tag) : "";
      const diff = cuoiNam - dauNam;
      const pct = dauNam !== 0 ? (diff / Math.abs(dauNam)) * 100 : cuoiNam !== 0 ? 100 : 0;

      const isHeader = ["110", "120", "130", "140", "150", "160", "170", "180", "200", "300", "400", "500"].includes(code);

      balanceSheet.push({
        maSo: code,
        tenChiTieu: DICT_CDKT_TT133[code] || `Chỉ tiêu ${code}`,
        thuyetMinh: tm,
        soCuoiNam: cuoiNam,
        soDauNam: dauNam,
        chenhLech: diff,
        phanTramThayDoi: Math.round(pct * 10) / 10,
        isHeader
      });
    }
  }

  // 3. Báo cáo Kết quả hoạt động kinh doanh (PL_KQHDSXKD)
  const profitLoss: BctcProfitLossItem[] = [];
  const plKq = xmlDoc.getElementsByTagName("PL_KQHDSXKD")[0];

  if (plKq) {
    const namNayEl = plKq.getElementsByTagName("NamNay")[0];
    const namTruocEl = plKq.getElementsByTagName("NamTruoc")[0];
    const thuyetMinhEl = plKq.getElementsByTagName("ThuyetMinh")[0];

    const codeSet = new Set<string>();
    if (namNayEl) {
      for (let i = 0; i < namNayEl.children.length; i++) {
        const tag = namNayEl.children[i].localName || namNayEl.children[i].tagName;
        if (tag.toLowerCase().startsWith("ct")) {
          codeSet.add(tag.toLowerCase().substring(2));
        }
      }
    }
    if (namTruocEl) {
      for (let i = 0; i < namTruocEl.children.length; i++) {
        const tag = namTruocEl.children[i].localName || namTruocEl.children[i].tagName;
        if (tag.toLowerCase().startsWith("ct")) {
          codeSet.add(tag.toLowerCase().substring(2));
        }
      }
    }

    const sortedCodes = Array.from(codeSet).sort((a, b) => parseInt(a, 10) - parseInt(b, 10));

    for (const code of sortedCodes) {
      const tag = `ct${code}`;
      const nn = namNayEl ? getNodeNumber(namNayEl, tag) : 0;
      const nt = namTruocEl ? getNodeNumber(namTruocEl, tag) : 0;
      const tm = thuyetMinhEl ? getNodeText(thuyetMinhEl, tag) : "";
      const diff = nn - nt;
      const pct = nt !== 0 ? (diff / Math.abs(nt)) * 100 : nn !== 0 ? 100 : 0;

      profitLoss.push({
        maSo: code,
        tenChiTieu: DICT_KQKD_TT133[code] || `Chỉ tiêu ${code}`,
        thuyetMinh: tm,
        namNay: nn,
        namTruoc: nt,
        chenhLech: diff,
        phanTramThayDoi: Math.round(pct * 10) / 10
      });
    }
  }

  // 4. Báo cáo Lưu chuyển tiền tệ (PL_LCTTTT hoặc PL_LCTTGT)
  const cashFlow: BctcCashFlowItem[] = [];
  const plLc = xmlDoc.getElementsByTagName("PL_LCTTTT")[0] || xmlDoc.getElementsByTagName("PL_LCTTGT")[0];

  if (plLc) {
    const namNayEl = plLc.getElementsByTagName("NamNay")[0];
    const namTruocEl = plLc.getElementsByTagName("NamTruoc")[0];
    const thuyetMinhEl = plLc.getElementsByTagName("ThuyetMinh")[0];

    const codeSet = new Set<string>();
    if (namNayEl) {
      for (let i = 0; i < namNayEl.children.length; i++) {
        const tag = namNayEl.children[i].localName || namNayEl.children[i].tagName;
        if (tag.toLowerCase().startsWith("ct")) {
          codeSet.add(tag.toLowerCase().substring(2));
        }
      }
    }

    const sortedCodes = Array.from(codeSet).sort((a, b) => parseInt(a, 10) - parseInt(b, 10));

    for (const code of sortedCodes) {
      const tag = `ct${code}`;
      const nn = namNayEl ? getNodeNumber(namNayEl, tag) : 0;
      const nt = namTruocEl ? getNodeNumber(namTruocEl, tag) : 0;
      const tm = thuyetMinhEl ? getNodeText(thuyetMinhEl, tag) : "";
      const diff = nn - nt;

      cashFlow.push({
        maSo: code,
        tenChiTieu: DICT_LCTT_TT133[code] || `Chỉ tiêu ${code}`,
        thuyetMinh: tm,
        namNay: nn,
        namTruoc: nt,
        chenhLech: diff
      });
    }
  }

  // 5. Bảng Cân đối tài khoản (PL_CDTK)
  const trialBalance: BctcTrialBalanceItem[] = [];
  const plCdtk = xmlDoc.getElementsByTagName("PL_CDTK")[0];

  let rawDuNoCuoi = 0;
  let rawDuCoCuoi = 0;

  if (plCdtk) {
    const soDuDau = plCdtk.getElementsByTagName("SoDuDauKy")[0] || plCdtk.getElementsByTagName("sodudauky")[0];
    const phatSinh =
      plCdtk.getElementsByTagName("SoPhatSinh")[0] ||
      plCdtk.getElementsByTagName("SoPhatSinhTrongKy")[0] ||
      plCdtk.getElementsByTagName("sophatsinh")[0];
    const soDuCuoi = plCdtk.getElementsByTagName("SoDuCuoiKy")[0] || plCdtk.getElementsByTagName("soducuoiky")[0];

    const dauNo = soDuDau ? soDuDau.getElementsByTagName("No")[0] : null;
    const dauCo = soDuDau ? soDuDau.getElementsByTagName("Co")[0] : null;

    const psNo = phatSinh ? phatSinh.getElementsByTagName("No")[0] : null;
    const psCo = phatSinh ? phatSinh.getElementsByTagName("Co")[0] : null;

    const cuoiNo = soDuCuoi ? soDuCuoi.getElementsByTagName("No")[0] : null;
    const cuoiCo = soDuCuoi ? soDuCuoi.getElementsByTagName("Co")[0] : null;

    // Thu thập tất cả tài khoản
    const accountSet = new Set<string>();
    const collectFromSection = (sec: Element | null) => {
      if (!sec) return;
      for (let i = 0; i < sec.children.length; i++) {
        const tag = sec.children[i].localName || sec.children[i].tagName;
        if (tag.toLowerCase().startsWith("ct") && tag.toLowerCase() !== "tongcong") {
          accountSet.add(tag.toLowerCase().substring(2));
        }
      }
    };

    collectFromSection(dauNo);
    collectFromSection(dauCo);
    collectFromSection(psNo);
    collectFromSection(psCo);
    collectFromSection(cuoiNo);
    collectFromSection(cuoiCo);

    // Lấy tổng cộng kiểm tra
    rawDuNoCuoi = cuoiNo ? getNodeNumber(cuoiNo, "tongCong") : 0;
    rawDuCoCuoi = cuoiCo ? getNodeNumber(cuoiCo, "tongCong") : 0;

    const sortedAccounts = Array.from(accountSet).sort();

    for (const tk of sortedAccounts) {
      const tag = `ct${tk}`;
      const dNo = dauNo ? getNodeNumber(dauNo, tag) : 0;
      const dCo = dauCo ? getNodeNumber(dauCo, tag) : 0;
      const pN = psNo ? getNodeNumber(psNo, tag) : 0;
      const pC = psCo ? getNodeNumber(psCo, tag) : 0;
      const cNo = cuoiNo ? getNodeNumber(cuoiNo, tag) : 0;
      const cCo = cuoiCo ? getNodeNumber(cuoiCo, tag) : 0;

      // Chỉ thêm nếu có phát sinh hoặc có số dư khác 0
      if (dNo !== 0 || dCo !== 0 || pN !== 0 || pC !== 0 || cNo !== 0 || cCo !== 0) {
        trialBalance.push({
          soHieuTK: tk,
          tenTK: DICT_TAI_KHOAN[tk] || `Tài khoản ${tk}`,
          duNoDau: dNo,
          duCoDau: dCo,
          psNo: pN,
          psCo: pC,
          duNoCuoi: cNo,
          duCoCuoi: cCo,
          capTK: tk.length <= 3 ? 1 : tk.length === 4 ? 2 : 3
        });
      }
    }
  }

  // 6. Tính toán tóm tắt & Các chỉ số tài chính (Financial Ratios)
  const getBsValue = (code: string, year: "cuoi" | "dau" = "cuoi"): number => {
    const item = balanceSheet.find(b => b.maSo === code);
    return item ? (year === "cuoi" ? item.soCuoiNam : item.soDauNam) : 0;
  };

  const getPlValue = (code: string): number => {
    const item = profitLoss.find(p => p.maSo === code);
    return item ? item.namNay : 0;
  };

  const tongTaiSanCuoi = getBsValue("200", "cuoi");
  const tongTaiSanDau = getBsValue("200", "dau");
  const noPhaiTraCuoi = getBsValue("300", "cuoi");
  const vonChuSoHuuCuoi = getBsValue("400", "cuoi");

  // Tài sản ngắn hạn (theo TT133 là tổng tiền + ĐTTC + Phải thu + Hàng tồn kho + TS khác)
  // Tiền: 110, ĐTTC: 120, Phải thu: 130, HTK: 140, TS khác: 180
  const tienCuoi = getBsValue("110", "cuoi");
  const phaiThuCuoi = getBsValue("130", "cuoi");
  const hangTonKhoCuoi = getBsValue("140", "cuoi");
  const taiSanNganHanCuoi = tienCuoi + getBsValue("120", "cuoi") + phaiThuCuoi + hangTonKhoCuoi + getBsValue("180", "cuoi");
  const noNganHanCuoi = noPhaiTraCuoi; // Trong TT133 đa số là nợ ngắn hạn hoặc tính trên mã 300

  const dtt = getPlValue("10") || getPlValue("01");
  const gvhb = getPlValue("11");
  const lnGop = getPlValue("20") || (dtt - gvhb);
  const lnTruocThue = getPlValue("50");
  const lnSauThue = getPlValue("60");

  const ratios: BctcFinancialRatios = {
    thanhToanHienHanh: noNganHanCuoi > 0 ? taiSanNganHanCuoi / noNganHanCuoi : 0,
    thanhToanNhanh: noNganHanCuoi > 0 ? (tienCuoi + phaiThuCuoi) / noNganHanCuoi : 0,
    tyLeTienMat: noNganHanCuoi > 0 ? tienCuoi / noNganHanCuoi : 0,
    heSoNoTrenTongTaiSan: tongTaiSanCuoi > 0 ? (noPhaiTraCuoi / tongTaiSanCuoi) * 100 : 0,
    heSoNoTrenVCSH: vonChuSoHuuCuoi > 0 ? (noPhaiTraCuoi / vonChuSoHuuCuoi) * 100 : 0,
    heSoTuTaiTro: tongTaiSanCuoi > 0 ? (vonChuSoHuuCuoi / tongTaiSanCuoi) * 100 : 0,
    bienLoiNhuanGop: dtt > 0 ? (lnGop / dtt) * 100 : 0,
    bienLoiNhuanRong: dtt > 0 ? (lnSauThue / dtt) * 100 : 0,
    roa: tongTaiSanCuoi > 0 ? (lnSauThue / tongTaiSanCuoi) * 100 : 0,
    roe: vonChuSoHuuCuoi > 0 ? (lnSauThue / vonChuSoHuuCuoi) * 100 : 0,
    vongQuayTongTaiSan: tongTaiSanCuoi > 0 ? dtt / tongTaiSanCuoi : 0,
    vongQuayHangTonKho: hangTonKhoCuoi > 0 ? gvhb / hangTonKhoCuoi : 0
  };

  const isTrialBalanceBalanced = Math.abs(rawDuNoCuoi - rawDuCoCuoi) <= 10; // Cho phép chênh lệch làm tròn nhỏ

  return {
    header,
    balanceSheet,
    profitLoss,
    cashFlow,
    trialBalance,
    ratios,
    rawSummary: {
      tongTaiSanCuoiNam: tongTaiSanCuoi,
      tongTaiSanDauNam: tongTaiSanDau,
      noPhaiTraCuoiNam: noPhaiTraCuoi,
      vonChuSoHuuCuoiNam: vonChuSoHuuCuoi,
      doanhThuThuan: dtt,
      giaVon: gvhb,
      loiNhuanGop: lnGop,
      loiNhuanTruocThue: lnTruocThue,
      loiNhuanSauThue: lnSauThue,
      tienVaTuongDuongTienCuoiKy: tienCuoi,
      tongDuNoCuoiKy: rawDuNoCuoi,
      tongDuCoCuoiKy: rawDuCoCuoi,
      isTrialBalanceBalanced
    }
  };
}

// Hàm xuất khẩu toàn bộ BCTC ra file Excel đa trang (Multi-sheet XLSX)
export function exportBctcToExcel(data: BctcParsedData, fileNamePrefix = "BCTC_BocTach"): void {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Thông tin chung
  const infoData = [
    ["HỆ THỐNG BÁO CÁO TÀI CHÍNH BÓC TÁCH TỪ XML THUẾ (HTKK)"],
    ["Tên doanh nghiệp", data.header.tenDoanhNghiep],
    ["Mã số thuế", data.header.mst],
    ["Địa chỉ", data.header.diaChi],
    ["Tỉnh/Thành phố", data.header.tinhTP],
    ["Mẫu tờ khai", `${data.header.tenTKhai} (Mã: ${data.header.maTKhai})`],
    ["Mô tả biểu mẫu", data.header.moTaBMau],
    ["Kỳ báo cáo", `Năm ${data.header.kyKKhai} (Từ ${data.header.kyKKhaiTuNgay} đến ${data.header.kyKKhaiDenNgay})`],
    ["Cơ quan thuế quản lý", `${data.header.tenCQT} (Mã: ${data.header.maCQT})`],
    ["Ngày lập tờ khai", data.header.ngayLapTKhai],
    ["Ngày ký số", data.header.ngayKy || "Chưa ký điện tử"],
    ["Kiểm toán", data.header.bctcDaKiemToan ? "Đã kiểm toán" : "Chưa kiểm toán"],
    ["Phiên bản XML", data.header.pbanTKhaiXML]
  ];
  const wsInfo = XLSX.utils.aoa_to_sheet(infoData);
  XLSX.utils.book_append_sheet(wb, wsInfo, "ThongTinChung");

  // Sheet 2: Bảng Cân đối kế toán
  const bsRows = [
    ["Mã số", "Chỉ tiêu", "Thuyết minh", "Số cuối năm (VNĐ)", "Số đầu năm (VNĐ)", "Chênh lệch (+/-)", "Tăng giảm (%)"]
  ];
  data.balanceSheet.forEach(row => {
    bsRows.push([
      row.maSo,
      row.tenChiTieu,
      row.thuyetMinh,
      row.soCuoiNam as any,
      row.soDauNam as any,
      row.chenhLech as any,
      row.phanTramThayDoi as any
    ]);
  });
  const wsBs = XLSX.utils.aoa_to_sheet(bsRows);
  XLSX.utils.book_append_sheet(wb, wsBs, "BangCanDoiKeToan");

  // Sheet 3: Kết quả hoạt động SXKD
  const plRows = [
    ["Mã số", "Chỉ tiêu", "Thuyết minh", "Năm nay (VNĐ)", "Năm trước (VNĐ)", "Chênh lệch (+/-)", "Tăng giảm (%)"]
  ];
  data.profitLoss.forEach(row => {
    plRows.push([
      row.maSo,
      row.tenChiTieu,
      row.thuyetMinh,
      row.namNay as any,
      row.namTruoc as any,
      row.chenhLech as any,
      row.phanTramThayDoi as any
    ]);
  });
  const wsPl = XLSX.utils.aoa_to_sheet(plRows);
  XLSX.utils.book_append_sheet(wb, wsPl, "KetQuaKinhDoanh");

  // Sheet 4: Lưu chuyển tiền tệ
  const cfRows = [
    ["Mã số", "Chỉ tiêu", "Thuyết minh", "Năm nay (VNĐ)", "Năm trước (VNĐ)", "Chênh lệch (+/-)"]
  ];
  data.cashFlow.forEach(row => {
    cfRows.push([
      row.maSo,
      row.tenChiTieu,
      row.thuyetMinh,
      row.namNay as any,
      row.namTruoc as any,
      row.chenhLech as any
    ]);
  });
  const wsCf = XLSX.utils.aoa_to_sheet(cfRows);
  XLSX.utils.book_append_sheet(wb, wsCf, "LuuChuyenTienTe");

  // Sheet 5: Bảng Cân đối tài khoản (Trial Balance)
  const tbRows = [
    ["Số hiệu TK", "Tên tài khoản", "Số dư Nợ đầu kỳ", "Số dư Có đầu kỳ", "Phát sinh Nợ", "Phát sinh Có", "Số dư Nợ cuối kỳ", "Số dư Có cuối kỳ"]
  ];
  data.trialBalance.forEach(row => {
    tbRows.push([
      row.soHieuTK,
      row.tenTK,
      row.duNoDau as any,
      row.duCoDau as any,
      row.psNo as any,
      row.psCo as any,
      row.duNoCuoi as any,
      row.duCoCuoi as any
    ]);
  });
  const wsTb = XLSX.utils.aoa_to_sheet(tbRows);
  XLSX.utils.book_append_sheet(wb, wsTb, "CanDoiTaiKhoan");

  // Sheet 6: Chỉ số tài chính
  const r = data.ratios;
  const ratioRows = [
    ["Nhóm chỉ số", "Tên chỉ số", "Giá trị", "Đơn vị", "Đánh giá sơ bộ"],
    ["Thanh khoản", "Khả năng thanh toán hiện hành (TSNH / Nợ NH)", r.thanhToanHienHanh.toFixed(2), "lần", r.thanhToanHienHanh >= 1.5 ? "Tốt" : r.thanhToanHienHanh >= 1 ? "Đạt" : "Rủi ro"],
    ["Thanh khoản", "Khả năng thanh toán nhanh", r.thanhToanNhanh.toFixed(2), "lần", r.thanhToanNhanh >= 1 ? "Tốt" : "Cần lưu ý"],
    ["Thanh khoản", "Tỷ số tiền mặt", r.tyLeTienMat.toFixed(2), "lần", r.tyLeTienMat >= 0.5 ? "Rất dồi dào" : "Bình thường"],
    ["Cơ cấu vốn", "Hệ số nợ trên Tổng tài sản (D/A)", r.heSoNoTrenTongTaiSan.toFixed(1), "%", r.heSoNoTrenTongTaiSan < 60 ? "An toàn" : "Đòn bẩy cao"],
    ["Cơ cấu vốn", "Hệ số nợ trên Vốn chủ sở hữu (D/E)", r.heSoNoTrenVCSH.toFixed(1), "%", r.heSoNoTrenVCSH < 150 ? "Bình thường" : "Cần thận trọng"],
    ["Cơ cấu vốn", "Hệ số tự tài trợ (VCSH / TTS)", r.heSoTuTaiTro.toFixed(1), "%", r.heSoTuTaiTro >= 40 ? "Độc lập tài chính cao" : "Phụ thuộc vốn vay"],
    ["Khả năng sinh lời", "Biên lợi nhuận gộp", r.bienLoiNhuanGop.toFixed(2), "%", "Hiệu quả sản xuất kinh doanh cốt lõi"],
    ["Khả năng sinh lời", "Biên lợi nhuận ròng (ROS)", r.bienLoiNhuanRong.toFixed(2), "%", "Tỷ lệ lãi ròng trên doanh thu"],
    ["Khả năng sinh lời", "Tỷ suất sinh lời trên Tổng tài sản (ROA)", r.roa.toFixed(2), "%", "Khả năng sinh lời của tài sản"],
    ["Khả năng sinh lời", "Tỷ suất sinh lời trên Vốn CSH (ROE)", r.roe.toFixed(2), "%", "Khả năng sinh lời cho cổ đông"],
    ["Hiệu quả hoạt động", "Vòng quay Tổng tài sản", r.vongQuayTongTaiSan.toFixed(2), "vòng/năm", "Tốc độ luân chuyển tài sản"],
    ["Hiệu quả hoạt động", "Vòng quay Hàng tồn kho", r.vongQuayHangTonKho.toFixed(2), "vòng/năm", "Tốc độ tiêu thụ hàng tồn kho"]
  ];
  const wsRatio = XLSX.utils.aoa_to_sheet(ratioRows);
  XLSX.utils.book_append_sheet(wb, wsRatio, "ChiSoTaiChinh");

  const safeFileName = `${fileNamePrefix}_${data.header.mst || "DoanhNghiep"}_Nam${data.header.kyKKhai || "2025"}.xlsx`;
  XLSX.writeFile(wb, safeFileName);
}

// Chuỗi XML mẫu thực tế mà người dùng vừa gửi (Công ty Cổ phần Đầu tư và Xây dựng Thương mại 69)
export const SAMPLE_BCTC_XML_STRING = `<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung>
      <TTinDVu>
        <maDVu>HTKK</maDVu>
        <tenDVu>HỖ TRỢ KÊ KHAI THUẾ</tenDVu>
        <pbanDVu>5.6.0</pbanDVu>
        <ttinNhaCCapDVu>93ACAAE1169A31A9DDE603DA37E71223</ttinNhaCCapDVu>
      </TTinDVu>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>683</maTKhai>
          <tenTKhai>Bộ báo cáo tài chính (B01a - DNN)(TT133/2016/TT-BTC)</tenTKhai>
          <moTaBMau>(Ban hành theo Thông tư số 133/2016/TT-BTC ngày 26/8/2016 của Bộ Tài chính)</moTaBMau>
          <pbanTKhaiXML>2.3.2</pbanTKhaiXML>
          <loaiTKhai>C</loaiTKhai>
          <soLan>0</soLan>
          <KyKKhaiThue>
            <kieuKy>Y</kieuKy>
            <kyKKhai>2025</kyKKhai>
            <kyKKhaiTuNgay>01/01/2025</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>31/12/2025</kyKKhaiDenNgay>
            <kyKKhaiTuThang />
            <kyKKhaiDenThang />
          </KyKKhaiThue>
          <maCQTNoiNop>11500</maCQTNoiNop>
          <tenCQTNoiNop>Thuế Tỉnh Hưng Yên</tenCQTNoiNop>
          <ngayLapTKhai>2026-03-18</ngayLapTKhai>
          <GiaHan>
            <maLyDoGiaHan>
            </maLyDoGiaHan>
            <lyDoGiaHan>
            </lyDoGiaHan>
          </GiaHan>
          <nguoiKy>
          </nguoiKy>
          <ngayKy>2026-03-18</ngayKy>
          <nganhNgheKD>
          </nganhNgheKD>
        </TKhaiThue>
        <NNT>
          <mst>0101437269</mst>
          <tenNNT>CÔNG TY CỔ PHẦN ĐẦU TƯ VÀ XÂY DỰNG THƯƠNG MẠI 69</tenNNT>
          <dchiNNT>Số 29 phố Nguyễn Thành, tổ dân phố số 05,</dchiNNT>
          <phuongXa>
          </phuongXa>
          <maHuyenNNT>
          </maHuyenNNT>
          <tenHuyenNNT>
          </tenHuyenNNT>
          <maTinhNNT>109</maTinhNNT>
          <tenTinhNNT>Tỉnh Hưng Yên</tenTinhNNT>
          <dthoaiNNT>
          </dthoaiNNT>
          <faxNNT>
          </faxNNT>
          <emailNNT>
          </emailNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <bctcDaKiemToan>0</bctcDaKiemToan>
      <maYKienKToan>0</maYKienKToan>
      <tenYKienKToan>
      </tenYKienKToan>
      <ThuyetMinh>
        <ct110></ct110>
        <ct120></ct120>
        <ct121></ct121>
        <ct122></ct122>
        <ct123></ct123>
        <ct124></ct124>
        <ct130></ct130>
        <ct131></ct131>
        <ct132></ct132>
        <ct133></ct133>
        <ct134></ct134>
        <ct135></ct135>
        <ct136></ct136>
        <ct140></ct140>
        <ct141></ct141>
        <ct142></ct142>
        <ct150></ct150>
        <ct151></ct151>
        <ct152></ct152>
        <ct160></ct160>
        <ct161></ct161>
        <ct162></ct162>
        <ct170></ct170>
        <ct180></ct180>
        <ct181></ct181>
        <ct182></ct182>
        <ct200></ct200>
        <ct300></ct300>
        <ct311></ct311>
        <ct312></ct312>
        <ct313></ct313>
        <ct314></ct314>
        <ct315></ct315>
        <ct316></ct316>
        <ct317></ct317>
        <ct318></ct318>
        <ct319></ct319>
        <ct320></ct320>
        <ct400></ct400>
        <ct411></ct411>
        <ct412></ct412>
        <ct413></ct413>
        <ct414></ct414>
        <ct415></ct415>
        <ct416></ct416>
        <ct417></ct417>
        <ct500></ct500>
      </ThuyetMinh>
      <SoCuoiNam>
        <ct110>17320616184</ct110>
        <ct120>0</ct120>
        <ct121>0</ct121>
        <ct122>0</ct122>
        <ct123>0</ct123>
        <ct124>0</ct124>
        <ct130>6540407245</ct130>
        <ct131>2165857200</ct131>
        <ct132>901558678</ct132>
        <ct133>0</ct133>
        <ct134>3472991367</ct134>
        <ct135>0</ct135>
        <ct136>0</ct136>
        <ct140>226000000</ct140>
        <ct141>226000000</ct141>
        <ct142>0</ct142>
        <ct150>848110905</ct150>
        <ct151>5159444647</ct151>
        <ct152>-4311333742</ct152>
        <ct160>0</ct160>
        <ct161>0</ct161>
        <ct162>0</ct162>
        <ct170>0</ct170>
        <ct180>44105488</ct180>
        <ct181>0</ct181>
        <ct182>44105488</ct182>
        <ct200>24979239822</ct200>
        <ct300>13146833893</ct300>
        <ct311>1936371647</ct311>
        <ct312>10957724078</ct312>
        <ct313>37991284</ct313>
        <ct314>132000000</ct314>
        <ct315>32259235</ct315>
        <ct316>0</ct316>
        <ct317>50487649</ct317>
        <ct318>0</ct318>
        <ct319>0</ct319>
        <ct320>0</ct320>
        <ct400>11832405929</ct400>
        <ct411>10000000000</ct411>
        <ct412>0</ct412>
        <ct413>0</ct413>
        <ct414>0</ct414>
        <ct415>0</ct415>
        <ct416>0</ct416>
        <ct417>1832405929</ct417>
        <ct500>24979239822</ct500>
      </SoCuoiNam>
      <SoDauNam>
        <ct110>5492794632</ct110>
        <ct120>1472650484</ct120>
        <ct121>0</ct121>
        <ct122>1472650484</ct122>
        <ct123>0</ct123>
        <ct124>0</ct124>
        <ct130>5621827008</ct130>
        <ct131>1792399600</ct131>
        <ct132>601658761</ct132>
        <ct133>0</ct133>
        <ct134>3227768647</ct134>
        <ct135>0</ct135>
        <ct136>0</ct136>
        <ct140>4308308320</ct140>
        <ct141>4308308320</ct141>
        <ct142>0</ct142>
        <ct150>1358618619</ct150>
        <ct151>5159444647</ct151>
        <ct152>-3800826028</ct152>
        <ct160>0</ct160>
        <ct161>0</ct161>
        <ct162>0</ct162>
        <ct170>0</ct170>
        <ct180>1098070625</ct180>
        <ct181>1079897898</ct181>
        <ct182>18172727</ct182>
        <ct200>19352269688</ct200>
        <ct300>7743341901</ct300>
        <ct311>2427548537</ct311>
        <ct312>5098785715</ct312>
        <ct313>0</ct313>
        <ct314>166520000</ct314>
        <ct315>0</ct315>
        <ct316>0</ct316>
        <ct317>50487649</ct317>
        <ct318>0</ct318>
        <ct319>0</ct319>
        <ct320>0</ct320>
        <ct400>11608927787</ct400>
        <ct411>10000000000</ct411>
        <ct412>0</ct412>
        <ct413>0</ct413>
        <ct414>0</ct414>
        <ct415>0</ct415>
        <ct416>0</ct416>
        <ct417>1608927787</ct417>
        <ct500>19352269688</ct500>
      </SoDauNam>
      <nguoiLapBieu></nguoiLapBieu>
      <keToanTruong></keToanTruong>
      <ngayLap>2026-03-16</ngayLap>
      <nguoiDaiDienTheoPhapLuat></nguoiDaiDienTheoPhapLuat>
    </CTieuTKhaiChinh>
    <PLuc>
      <PL_KQHDSXKD>
        <ThuyetMinh>
          <ct01 /><ct02 /><ct10 /><ct11 /><ct20 /><ct21 /><ct22 /><ct23 /><ct24 /><ct30 /><ct31 /><ct32 /><ct40 /><ct50 /><ct51 /><ct60 />
        </ThuyetMinh>
        <NamNay>
          <ct01>18704078451</ct01>
          <ct02>0</ct02>
          <ct10>18704078451</ct10>
          <ct11>15871971238</ct11>
          <ct20>2832107213</ct20>
          <ct21>3874307</ct21>
          <ct22>0</ct22>
          <ct23>0</ct23>
          <ct24>2614677378</ct24>
          <ct30>221304142</ct30>
          <ct31>2174000</ct31>
          <ct32>0</ct32>
          <ct40>2174000</ct40>
          <ct50>223478142</ct50>
          <ct51>37991284</ct51>
          <ct60>185486858</ct60>
        </NamNay>
        <NamTruoc>
          <ct01>47307851836</ct01>
          <ct02>0</ct02>
          <ct10>47307851836</ct10>
          <ct11>43706675627</ct11>
          <ct20>3601176209</ct20>
          <ct21>325842755</ct21>
          <ct22>0</ct22>
          <ct23>0</ct23>
          <ct24>3176875216</ct24>
          <ct30>750143748</ct30>
          <ct31>5838437</ct31>
          <ct32>13216</ct32>
          <ct40>5825221</ct40>
          <ct50>755968969</ct50>
          <ct51>151275897</ct51>
          <ct60>604693072</ct60>
        </NamTruoc>
      </PL_KQHDSXKD>
      <PL_LCTTTT>
        <ThuyetMinh>
          <ct01 /><ct02 /><ct03 /><ct04 /><ct05 /><ct06 /><ct07 /><ct20 /><ct21 /><ct22 /><ct23 /><ct24 /><ct25 /><ct30 /><ct31 /><ct32 /><ct33 /><ct34 /><ct35 /><ct40 /><ct50 /><ct60 /><ct61 /><ct70 />
        </ThuyetMinh>
        <NamNay>
          <ct01>25941836872</ct01>
          <ct02>-12656179389</ct02>
          <ct03>-1570720000</ct03>
          <ct04>0</ct04>
          <ct05>0</ct05>
          <ct06>3073514054</ct06>
          <ct07>-4433280469</ct07>
          <ct20>10355171068</ct20>
          <ct21>0</ct21>
          <ct22>0</ct22>
          <ct23>0</ct23>
          <ct24>1472650484</ct24>
          <ct25>0</ct25>
          <ct30>1472650484</ct30>
          <ct31>0</ct31>
          <ct32>0</ct32>
          <ct33>0</ct33>
          <ct34>0</ct34>
          <ct35>0</ct35>
          <ct40>0</ct40>
          <ct50>11827821552</ct50>
          <ct60>5492794632</ct60>
          <ct61>0</ct61>
          <ct70>17320616184</ct70>
        </NamNay>
        <NamTruoc>
          <ct01>47509531271</ct01>
          <ct02>-45424453838</ct02>
          <ct03>-1932550000</ct03>
          <ct04>0</ct04>
          <ct05>-152000000</ct05>
          <ct06>8650256436</ct06>
          <ct07>-8173754814</ct07>
          <ct20>477029055</ct20>
          <ct21>0</ct21>
          <ct22>0</ct22>
          <ct23>0</ct23>
          <ct24>3000000000</ct24>
          <ct25>0</ct25>
          <ct30>3000000000</ct30>
          <ct31>0</ct31>
          <ct32>0</ct32>
          <ct33>0</ct33>
          <ct34>0</ct34>
          <ct35>0</ct35>
          <ct40>0</ct40>
          <ct50>3477029055</ct50>
          <ct60>2015765577</ct60>
          <ct61>0</ct61>
          <ct70>5492794632</ct70>
        </NamTruoc>
      </PL_LCTTTT>
      <PL_CDTK>
        <SoDuDauKy>
          <No>
            <ct111>485259392</ct111>
            <ct1111>485259392</ct1111>
            <ct1112>0</ct1112>
            <ct112>5007535240</ct112>
            <ct1121>5007535240</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>1472650484</ct128>
            <ct1281>0</ct1281>
            <ct1288>1472650484</ct1288>
            <ct131>1792399600</ct131>
            <ct133>1079897898</ct133>
            <ct1331>1079897898</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>1368774869</ct138>
            <ct1381>0</ct1381>
            <ct1386>307726200</ct1386>
            <ct1388>1061048669</ct1388>
            <ct141>1965537</ct141>
            <ct151>0</ct151>
            <ct152>0</ct152>
            <ct153>0</ct153>
            <ct154>4308308320</ct154>
            <ct155>0</ct155>
            <ct156>0</ct156>
            <ct157>0</ct157>
            <ct211>5159444647</ct211>
            <ct2111>5159444647</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>0</ct214>
            <ct2141>0</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>18172727</ct242>
            <ct331>601658761</ct331>
            <ct333>0</ct333>
            <ct3331>0</ct3331>
            <ct33311>0</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>0</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>0</ct3339>
            <ct334>0</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>1889287476</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>0</ct3383>
            <ct3384>0</ct3384>
            <ct3385>0</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>1889287476</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>0</ct411>
            <ct4111>0</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>0</ct421>
            <ct4211>0</ct4211>
            <ct4212>0</ct4212>
            <ct511>0</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>0</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>0</ct632>
            <ct635>0</ct635>
            <ct642>0</ct642>
            <ct6421>0</ct6421>
            <ct6422>0</ct6422>
            <ct711>0</ct711>
            <ct811>0</ct811>
            <ct821>0</ct821>
            <ct911>0</ct911>
            <tongCong>23185354951</tongCong>
          </No>
          <Co>
            <ct111>0</ct111>
            <ct1111>0</ct1111>
            <ct1112>0</ct1112>
            <ct112>0</ct112>
            <ct1121>0</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>5098785715</ct131>
            <ct133>0</ct133>
            <ct1331>0</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>0</ct138>
            <ct1381>0</ct1381>
            <ct1386>0</ct1386>
            <ct1388>0</ct1388>
            <ct141>0</ct141>
            <ct151>0</ct151>
            <ct152>0</ct152>
            <ct153>0</ct153>
            <ct154>0</ct154>
            <ct155>0</ct155>
            <ct156>0</ct156>
            <ct157>0</ct157>
            <ct211>0</ct211>
            <ct2111>0</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>3800826028</ct214>
            <ct2141>3800826028</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>0</ct242>
            <ct331>2427548537</ct331>
            <ct333>0</ct333>
            <ct3331>0</ct3331>
            <ct33311>0</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>0</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>0</ct3339>
            <ct334>166520000</ct334>
            <ct335>0</ct335>
            <ct336>50487649</ct336>
            <ct3361>0</ct3361>
            <ct3368>50487649</ct3368>
            <ct338>32259235</ct338>
            <ct3381>7972699</ct3381>
            <ct3382>24286536</ct3382>
            <ct3383>0</ct3383>
            <ct3384>0</ct3384>
            <ct3385>0</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>0</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>10000000000</ct411>
            <ct4111>10000000000</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>1608927787</ct421>
            <ct4211>1004234715</ct4211>
            <ct4212>604693072</ct4212>
            <ct511>0</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>0</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>0</ct632>
            <ct635>0</ct635>
            <ct642>0</ct642>
            <ct6421>0</ct6421>
            <ct6422>0</ct6422>
            <ct711>0</ct711>
            <ct811>0</ct811>
            <ct821>0</ct821>
            <ct911>0</ct911>
            <tongCong>23185354951</tongCong>
          </Co>
        </SoDuDauKy>
        <SoPhatSinhTrongKy>
          <No>
            <ct111>2060000000</ct111>
            <ct1111>2060000000</ct1111>
            <ct1112>0</ct1112>
            <ct112>30849001410</ct112>
            <ct1121>30849001410</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>21125219237</ct131>
            <ct133>930222478</ct133>
            <ct1331>930222478</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>1959730057</ct138>
            <ct1381>0</ct1381>
            <ct1386>7463435</ct1386>
            <ct1388>1952266622</ct1388>
            <ct141>30000000</ct141>
            <ct151>0</ct151>
            <ct152>7374103336</ct152>
            <ct153>0</ct153>
            <ct154>11563662918</ct154>
            <ct155>0</ct155>
            <ct156>0</ct156>
            <ct157>0</ct157>
            <ct211>0</ct211>
            <ct2111>0</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>0</ct214>
            <ct2141>0</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>0</ct242>
            <ct331>13517319288</ct331>
            <ct333>1498326277</ct333>
            <ct3331>1496326277</ct3331>
            <ct33311>1496326277</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>0</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>2000000</ct3339>
            <ct334>1657687500</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>387055000</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>202528000</ct3383>
            <ct3384>35757400</ct3384>
            <ct3385>15189600</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>133580000</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>0</ct411>
            <ct4111>0</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>604693072</ct421>
            <ct4211>0</ct4211>
            <ct4212>604693072</ct4212>
            <ct511>18704078451</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>18704078451</ct5118>
            <ct515>3874307</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>15871971238</ct632>
            <ct635>0</ct635>
            <ct642>2614677378</ct642>
            <ct6421>0</ct6421>
            <ct6422>2614677378</ct6422>
            <ct711>2174000</ct711>
            <ct811>0</ct811>
            <ct821>37991284</ct821>
            <ct911>18710126758</ct911>
            <tongCong>149501913989</tongCong>
          </No>
          <Co>
            <ct111>1574372434</ct111>
            <ct1111>1574372434</ct1111>
            <ct1112>0</ct1112>
            <ct112>19506807424</ct112>
            <ct1121>19506807424</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>1472650484</ct128>
            <ct1281>0</ct1281>
            <ct1288>1472650484</ct1288>
            <ct131>26610700000</ct131>
            <ct133>2010120376</ct133>
            <ct1331>2010120376</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>1332202287</ct138>
            <ct1381>0</ct1381>
            <ct1386>230160054</ct1386>
            <ct1388>1102042233</ct1388>
            <ct141>6032776</ct141>
            <ct151>0</ct151>
            <ct152>7148103336</ct152>
            <ct153>0</ct153>
            <ct154>15871971238</ct154>
            <ct155>0</ct155>
            <ct156>0</ct156>
            <ct157>0</ct157>
            <ct211>0</ct211>
            <ct2111>0</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>510507714</ct214>
            <ct2141>510507714</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>0</ct242>
            <ct331>12726242481</ct331>
            <ct333>1536317561</ct333>
            <ct3331>1496326277</ct3331>
            <ct33311>1496326277</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>37991284</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>2000000</ct3339>
            <ct334>1623167500</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>799653748</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>187264000</ct3383>
            <ct3384>33086200</ct3384>
            <ct3385>14044800</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>565258748</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>0</ct411>
            <ct4111>0</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>828171214</ct421>
            <ct4211>604693072</ct4211>
            <ct4212>223478142</ct4212>
            <ct511>18704078451</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>18704078451</ct5118>
            <ct515>3874307</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>15871971238</ct632>
            <ct635>0</ct635>
            <ct642>2614677378</ct642>
            <ct6421>0</ct6421>
            <ct6422>2614677378</ct6422>
            <ct711>2174000</ct711>
            <ct811>0</ct811>
            <ct821>37991284</ct821>
            <ct911>18710126758</ct911>
            <tongCong>149501913989</tongCong>
          </Co>
        </SoPhatSinhTrongKy>
        <SoDuCuoiKy>
          <No>
            <ct111>970886958</ct111>
            <ct1111>970886958</ct1111>
            <ct1112>0</ct1112>
            <ct112>16349729226</ct112>
            <ct1121>16349729226</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>2165857200</ct131>
            <ct133>0</ct133>
            <ct1331>0</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>1996302637</ct138>
            <ct1381>0</ct1381>
            <ct1386>85029581</ct1386>
            <ct1388>1911273056</ct1388>
            <ct141>25932761</ct141>
            <ct151>0</ct151>
            <ct152>226000000</ct152>
            <ct153>0</ct153>
            <ct154>0</ct154>
            <ct155>0</ct155>
            <ct156>0</ct156>
            <ct157>0</ct157>
            <ct211>5159444647</ct211>
            <ct2111>5159444647</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>0</ct214>
            <ct2141>0</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>18172727</ct242>
            <ct331>901558678</ct331>
            <ct333>0</ct333>
            <ct3331>0</ct3331>
            <ct33311>0</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>0</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>0</ct3339>
            <ct334>0</ct334>
            <ct335>0</ct335>
            <ct336>0</ct336>
            <ct3361>0</ct3361>
            <ct3368>0</ct3368>
            <ct338>1476688728</ct338>
            <ct3381>0</ct3381>
            <ct3382>0</ct3382>
            <ct3383>15264000</ct3383>
            <ct3384>2671200</ct3384>
            <ct3385>1144800</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>1457608728</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>0</ct411>
            <ct4111>0</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>0</ct421>
            <ct4211>0</ct4211>
            <ct4212>0</ct4212>
            <ct511>0</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>0</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>0</ct632>
            <ct635>0</ct635>
            <ct642>0</ct642>
            <ct6421>0</ct6421>
            <ct6422>0</ct6422>
            <ct711>0</ct711>
            <ct811>0</ct811>
            <ct821>0</ct821>
            <ct911>0</ct911>
            <tongCong>29290573562</tongCong>
          </No>
          <Co>
            <ct111>0</ct111>
            <ct1111>0</ct1111>
            <ct1112>0</ct1112>
            <ct112>0</ct112>
            <ct1121>0</ct1121>
            <ct1122>0</ct1122>
            <ct121>0</ct121>
            <ct128>0</ct128>
            <ct1281>0</ct1281>
            <ct1288>0</ct1288>
            <ct131>10957724078</ct131>
            <ct133>0</ct133>
            <ct1331>0</ct1331>
            <ct1332>0</ct1332>
            <ct136>0</ct136>
            <ct1361>0</ct1361>
            <ct1368>0</ct1368>
            <ct138>0</ct138>
            <ct1381>0</ct1381>
            <ct1386>0</ct1386>
            <ct1388>0</ct1388>
            <ct141>0</ct141>
            <ct151>0</ct151>
            <ct152>0</ct152>
            <ct153>0</ct153>
            <ct154>0</ct154>
            <ct155>0</ct155>
            <ct156>0</ct156>
            <ct157>0</ct157>
            <ct211>0</ct211>
            <ct2111>0</ct2111>
            <ct2112>0</ct2112>
            <ct2113>0</ct2113>
            <ct214>4311333742</ct214>
            <ct2141>4311333742</ct2141>
            <ct2142>0</ct2142>
            <ct2143>0</ct2143>
            <ct2147>0</ct2147>
            <ct217>0</ct217>
            <ct228>0</ct228>
            <ct2281>0</ct2281>
            <ct2288>0</ct2288>
            <ct229>0</ct229>
            <ct2291>0</ct2291>
            <ct2292>0</ct2292>
            <ct2293>0</ct2293>
            <ct2294>0</ct2294>
            <ct241>0</ct241>
            <ct2411>0</ct2411>
            <ct2412>0</ct2412>
            <ct2413>0</ct2413>
            <ct242>0</ct242>
            <ct331>1936371647</ct331>
            <ct333>37991284</ct333>
            <ct3331>0</ct3331>
            <ct33311>0</ct33311>
            <ct33312>0</ct33312>
            <ct3332>0</ct3332>
            <ct3333>0</ct3333>
            <ct3334>37991284</ct3334>
            <ct3335>0</ct3335>
            <ct3336>0</ct3336>
            <ct3337>0</ct3337>
            <ct3338>0</ct3338>
            <ct33381>0</ct33381>
            <ct33382>0</ct33382>
            <ct3339>0</ct3339>
            <ct334>132000000</ct334>
            <ct335>0</ct335>
            <ct336>50487649</ct336>
            <ct3361>0</ct3361>
            <ct3368>50487649</ct3368>
            <ct338>32259235</ct338>
            <ct3381>7972699</ct3381>
            <ct3382>24286536</ct3382>
            <ct3383>0</ct3383>
            <ct3384>0</ct3384>
            <ct3385>0</ct3385>
            <ct3386>0</ct3386>
            <ct3387>0</ct3387>
            <ct3388>0</ct3388>
            <ct341>0</ct341>
            <ct3411>0</ct3411>
            <ct3412>0</ct3412>
            <ct352>0</ct352>
            <ct3521>0</ct3521>
            <ct3522>0</ct3522>
            <ct3524>0</ct3524>
            <ct353>0</ct353>
            <ct3531>0</ct3531>
            <ct3532>0</ct3532>
            <ct3533>0</ct3533>
            <ct3534>0</ct3534>
            <ct356>0</ct356>
            <ct3561>0</ct3561>
            <ct3562>0</ct3562>
            <ct411>10000000000</ct411>
            <ct4111>10000000000</ct4111>
            <ct4112>0</ct4112>
            <ct4118>0</ct4118>
            <ct413>0</ct413>
            <ct418>0</ct418>
            <ct419>0</ct419>
            <ct421>1832405929</ct421>
            <ct4211>1608927787</ct4211>
            <ct4212>223478142</ct4212>
            <ct511>0</ct511>
            <ct5111>0</ct5111>
            <ct5112>0</ct5112>
            <ct5113>0</ct5113>
            <ct5118>0</ct5118>
            <ct515>0</ct515>
            <ct611>0</ct611>
            <ct631>0</ct631>
            <ct632>0</ct632>
            <ct635>0</ct635>
            <ct642>0</ct642>
            <ct6421>0</ct6421>
            <ct6422>0</ct6422>
            <ct711>0</ct711>
            <ct811>0</ct811>
            <ct821>0</ct821>
            <ct911>0</ct911>
            <tongCong>29290573564</tongCong>
          </Co>
        </SoDuCuoiKy>
      </PL_CDTK>
    </PLuc>
  </HSoKhaiThue>
</HSoThueDTu>`;
