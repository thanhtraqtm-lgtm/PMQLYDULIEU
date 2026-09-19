// Tiện ích phân tích Báo cáo tài chính (XML / PDF / Excel) và Đối chiếu Mẫu Phiếu 01/IO-DN
// Tuân thủ Hướng dẫn Thông tư 200/2014/TT-BTC & Thông tư 133/2016/TT-BTC của Tổng cục Thống kê (IO-2026)
import * as XLSX from "xlsx";

export interface RawBctcData {
  mst: string;
  tenDoanhNghiep: string;
  diaChi: string;
  nam: number;
  thongTu: "TT200" | "TT133" | "Khac";
  // Kết quả kinh doanh
  dtt_10: number; // Doanh thu thuần (Mã 10 / TK 511)
  gvhb_11: number; // Giá vốn hàng bán (Mã 11 / TK 632)
  cfbh_25: number; // Chi phí bán hàng (Mã 25 / TK 641 / TK 6421)
  cfql_26: number; // Chi phí quản lý (Mã 26 / TK 642 / TK 6422)
  laiVay_23: number; // Chi phí lãi vay (Mã 23 / TK 635)
  ln_thuan_30: number; // Lợi nhuận thuần (Mã 30)
  ln_truoc_thue_50: number; // Lợi nhuận trước thuế (Mã 50)
  // Bảng cân đối kế toán & Tài khoản tồn kho
  tonKho_141: number; // Hàng tồn kho (Mã 141)
  cfsxdd_154_dauKy: number; // Số dư Nợ đầu kỳ TK 154
  cfsxdd_154_cuoiKy: number; // Số dư Nợ cuối kỳ TK 154
  thanhPham_155_dauKy: number; // Số dư Nợ đầu kỳ TK 155
  thanhPham_155_cuoiKy: number; // Số dư Nợ cuối kỳ TK 155
  hangGuiBan_157_dauKy: number; // Số dư Nợ đầu kỳ TK 157
  hangGuiBan_157_cuoiKy: number; // Số dư Nợ cuối kỳ TK 157
  hangHoa_156_dauKy: number; // Số dư Nợ đầu kỳ TK 156
  hangHoa_156_cuoiKy: number; // Số dư Nợ cuối kỳ TK 156
  // Chi phí sản xuất (Số phát sinh)
  cf_nvl_621: number; // Nợ TK 621 hoặc TK 154 ứng Có 152
  cf_nc_622: number; // Nợ TK 622 hoặc TK 154 ứng Có 334
  cf_sxc_627: number; // Nợ TK 627 hoặc TK 154 ứng Có khác
  cf_mtc_623: number; // Nợ TK 623 hoặc TK 154 máy thi công
  // Khấu hao & Thuế & Nhân công
  khauHao_214: number; // Khấu hao TSCĐ (Có TK 214)
  thue_333: number; // Thuế phát sinh phải nộp (TK 333 trừ 3334, 3335)
  tienLuong_334: number; // Tiền lương phải trả NLĐ (Có TK 334)
  bhxh_338: number; // BHXH, BHYT, KPCĐ (Có TK 338)
  // Nợ vay
  vayNganHan_320: number; // Mã số 320 CĐKT
  vayDaiHan_338: number; // Mã số 338 CĐKT
  // Metadata nguồn
  sourceFileName: string;
  sourceType: "xml" | "pdf" | "excel";
}

export interface Phieu01IODNForm {
  thongTinDinhDanh: {
    tenDoanhNghiep: string;
    maSoThue: string;
    diaChi: string;
    tinhTP: string;
    soDienThoai: string;
    loaiHinhKinhTe: string; // 1: Nhà nước, 2: Ngoài nhà nước, 3: FDI
    maNganhC5: string;
    maIO: string;
    tenNganhChinh: string;
    sectorId?: string;
    sectorExplanation?: string;
  };
  phan1: {
    cau5_tongDoanhThuThuan: number;
    cau6_doanhThuSanPham: Array<{
      stt: number;
      tenNganh: string;
      maIO: string;
      c1_dtt: number;
      c2_coGiaCong: number; // 1: Có, 2: Không
      c3_phiGiaCong: number;
    }>;
    cau7_thongTinKhac: Array<{
      stt: number;
      tenNganh: string;
      maIO: string;
      c1_gvhb: number;
      c2_gvChuyenBan: number;
      c3_chiHo: number;
      c4_chiThuong: number;
    }>;
    cau8_hangTonKhoDN: {
      c81_sxdd_dauKy: number;
      c81_sxdd_cuoiKy: number;
      c82_thanhPham_dauKy: number;
      c82_thanhPham_cuoiKy: number;
      c83_hangGui_dauKy: number;
      c83_hangGui_cuoiKy: number;
      tong_dauKy: number;
      tong_cuoiKy: number;
    };
    cau9_tonKhoSanPham: Array<{
      stt: number;
      tenNganh: string;
      maIO: string;
      c1_sxdd_dau: number;
      c2_sxdd_cuoi: number;
      c3_tp_dau: number;
      c4_tp_cuoi: number;
      c5_gui_dau: number;
      c6_gui_cuoi: number;
    }>;
  };
  phan2: {
    cau10_sxkdChinh: {
      dtt: number;
      gvhb: number;
      cfbh: number;
      cfql: number;
      ma223_loiNhuan: number;
      ma224_laiVay: number;
      cfnvl: number;
      cfnc: number;
      cfsxc: number;
      cfmtc: number;
      // Cột 2 (Gia công)
      c2_dtt: number;
      c2_gvhb: number;
      c2_cfbh: number;
      c2_cfql: number;
      c2_loiNhuan: number;
      c2_laiVay: number;
      c2_cfnvl: number;
      c2_cfnc: number;
      c2_cfsxc: number;
      c2_cfmtc: number;
    };
    cau11_troCap: Array<{ maIO: string; ten: string; soTien: number }>;
    cau12_thueLePhi: Array<{ maIO: string; ten: string; soTien: number; duocHachToan: boolean }>;
    cau13_tienThuong191: number;
    cau14_thuongMai: {
      c1_dtt: number;
      c2_muaVe: number;
      c3_vanTai: number;
      c4_tonKhoDau: number;
      c5_tonKhoCuoi: number;
      c6_phiTM: number;
    };
  };
  phan3: {
    cau15_nvlChiTiet: Array<{ maIO: string; moTa: string; c1_giaTri: number; c2_tyLeNK: number; c3_giaCong: number }>;
    cau16_nhanCong: {
      ma182_tienLuong: number;
      ma183_bhxh: number;
      ma184_bhyt: number;
      ma185_bhtn: number;
      ma186_bhConNguoi: number;
      ma187_kpcd: number;
      ma195_chiTrucTiep: number;
      ma196_chiTraKhac: number;
      tongCong: number;
      // Cột gia công
      gc_182: number;
      gc_183: number;
      gc_184: number;
      gc_185: number;
      gc_186: number;
      gc_187: number;
      gc_195: number;
      gc_196: number;
      gc_tong: number;
    };
    cau17_khauHao225: number;
    cau18_muaNgoaiChiPhiKhac: Array<{ maIO: string; moTa: string; c1_giaTri: number; c2_tyLeNK: number }>;
  };
  phan4: {
    cau194_vayNganHan320: number;
    cau194_vayDaiHan338: number;
    cau191_coTroCap: boolean;
    cau192_coBaoLanh: boolean;
    cau193_coVayNN: boolean;
  };
  // Chỉ số kinh tế tổng hợp IO
  tongHopIO: {
    ic: number;
    go: number;
    tyLeIC_GO: number;
    danhGiaIC: "bình thường" | "quá cao (>0.75 hoặc >0.85)" | "quá thấp (<0.3)";
  };
}

// 1. Phân tích nội dung XML BCTC (HTKK / eTax / Fast / MISA)
export function parseBctcXml(xmlText: string, fileName: string): RawBctcData {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlText, "text/xml");

  const parseNum = (text: string | null | undefined): number => {
    if (!text) return 0;
    const clean = text.replace(/,/g, "").replace(/\s/g, "");
    const n = Number(clean);
    return isNaN(n) ? 0 : n;
  };

  const getTagText = (tagNames: string[]): string => {
    for (const tag of tagNames) {
      const el = xmlDoc.getElementsByTagName(tag)[0];
      if (el && el.textContent) return el.textContent.trim();

      const lower = tag.toLowerCase();
      const allEls = xmlDoc.getElementsByTagName("*");
      for (let i = 0; i < allEls.length; i++) {
        if (allEls[i].localName.toLowerCase() === lower && allEls[i].textContent) {
          return allEls[i].textContent.trim();
        }
      }
    }
    return "";
  };

  // Trích xuất thông tin doanh nghiệp
  const mst = getTagText(["mst", "MaSoThue", "MST", "tin"]);
  const tenDoanhNghiep = getTagText(["tenNNT", "TenDoanhNghiep", "TenNNT", "ten_nnt", "ten"]);
  const diaChi = getTagText(["dchiNNT", "DiaChi", "DChiNNT", "dia_chi", "address"]);
  const namStr = getTagText(["kyKKhai", "nam", "NamBCTC", "nam_bctc", "year"]);
  const nam = parseInt(namStr) || 2025;

  // Kiểm tra thông tư áp dụng
  const fullContent = xmlText.toLowerCase();
  let thongTu: "TT200" | "TT133" | "Khac" = "TT200";
  if (
    fullContent.includes("133/2016") ||
    fullContent.includes("tt133") ||
    fullContent.includes("b01a - dnn") ||
    fullContent.includes("b01b - dnn")
  ) {
    thongTu = "TT133";
  }

  // 1. Helper tìm giá trị chỉ tiêu từ Kết quả kinh doanh (PL_KQHDSXKD)
  const getKqkdValue = (code: string): number => {
    const plKq = xmlDoc.getElementsByTagName("PL_KQHDSXKD")[0] ||
                 xmlDoc.getElementsByTagName("pl_kqhdsxkd")[0];
    if (plKq) {
      const namNay = plKq.getElementsByTagName("NamNay")[0] ||
                     plKq.getElementsByTagName("namnay")[0];
      if (namNay) {
        for (let i = 0; i < namNay.children.length; i++) {
          const child = namNay.children[i];
          const name = child.localName.toLowerCase();
          if (name === `ct${code}` || name === `ct_${code}` || name === `ms${code}`) {
            const v = parseNum(child.textContent);
            if (v !== 0) return v;
          }
        }
      }
    }

    // Fallback tìm thẻ dạng ct{code} hoặc có thuộc tính maSo="{code}"
    const allEls = xmlDoc.getElementsByTagName("*");
    for (let i = 0; i < allEls.length; i++) {
      const el = allEls[i];
      const maSoAttr = el.getAttribute("maSo") || el.getAttribute("MaSo") || el.getAttribute("code");
      if (maSoAttr === code) {
        const valText = el.getAttribute("soNamNay") || el.getAttribute("soCuoiNam") || el.getAttribute("giaTri") || el.textContent;
        const v = parseNum(valText);
        if (v !== 0) return v;
      }
      const local = el.localName.toLowerCase();
      if (local === `ct${code}` || local === `ct_${code}` || local === `ms${code}`) {
        if (!el.parentElement || el.parentElement.localName.toLowerCase() !== "thuyetminh") {
          const v = parseNum(el.textContent);
          if (v !== 0) return v;
        }
      }
    }
    return 0;
  };

  // 2. Helper lấy giá trị từ Bảng cân đối kế toán (CTieuTKhaiChinh > SoCuoiNam / SoDauNam)
  const getCdktValue = (code: string, period: "cuoiNam" | "dauNam" = "cuoiNam"): number => {
    const parentTag = period === "cuoiNam" ? "SoCuoiNam" : "SoDauNam";
    const parentEl = xmlDoc.getElementsByTagName(parentTag)[0] ||
                     xmlDoc.getElementsByTagName(parentTag.toLowerCase())[0];
    if (parentEl) {
      for (let i = 0; i < parentEl.children.length; i++) {
        const child = parentEl.children[i];
        const name = child.localName.toLowerCase();
        if (name === `ct${code}` || name === `ct_${code}`) {
          return parseNum(child.textContent);
        }
      }
    }
    return 0;
  };

  // 3. Helper lấy số dư / phát sinh từ Bảng Cân đối tài khoản (PL_CDTK)
  const getCdtkAccountAmount = (
    tkPrefix: string,
    field: "duNoDau" | "duCoDau" | "psNo" | "psCo" | "duNoCuoi" | "duCoCuoi"
  ): number => {
    // A. Cấu trúc chuẩn HTKK: <PL_CDTK>
    const plCdtk = xmlDoc.getElementsByTagName("PL_CDTK")[0] ||
                   xmlDoc.getElementsByTagName("pl_cdtk")[0];
    if (plCdtk) {
      let targetSection: Element | null = null;
      if (field === "duNoDau") {
        const dauKy = plCdtk.getElementsByTagName("SoDuDauKy")[0] || plCdtk.getElementsByTagName("sodudauky")[0];
        targetSection = dauKy ? (dauKy.getElementsByTagName("No")[0] || dauKy.getElementsByTagName("no")[0]) : null;
      } else if (field === "duCoDau") {
        const dauKy = plCdtk.getElementsByTagName("SoDuDauKy")[0] || plCdtk.getElementsByTagName("sodudauky")[0];
        targetSection = dauKy ? (dauKy.getElementsByTagName("Co")[0] || dauKy.getElementsByTagName("co")[0]) : null;
      } else if (field === "psNo") {
        const ps = plCdtk.getElementsByTagName("SoPhatSinh")[0] || plCdtk.getElementsByTagName("SoPhatSinhTrongKy")[0] || plCdtk.getElementsByTagName("sophatsinh")[0];
        targetSection = ps ? (ps.getElementsByTagName("No")[0] || ps.getElementsByTagName("no")[0]) : null;
      } else if (field === "psCo") {
        const ps = plCdtk.getElementsByTagName("SoPhatSinh")[0] || plCdtk.getElementsByTagName("SoPhatSinhTrongKy")[0] || plCdtk.getElementsByTagName("sophatsinh")[0];
        targetSection = ps ? (ps.getElementsByTagName("Co")[0] || ps.getElementsByTagName("co")[0]) : null;
      } else if (field === "duNoCuoi") {
        const cuoiKy = plCdtk.getElementsByTagName("SoDuCuoiKy")[0] || plCdtk.getElementsByTagName("soducuoiky")[0];
        targetSection = cuoiKy ? (cuoiKy.getElementsByTagName("No")[0] || cuoiKy.getElementsByTagName("no")[0]) : null;
      } else if (field === "duCoCuoi") {
        const cuoiKy = plCdtk.getElementsByTagName("SoDuCuoiKy")[0] || plCdtk.getElementsByTagName("soducuoiky")[0];
        targetSection = cuoiKy ? (cuoiKy.getElementsByTagName("Co")[0] || cuoiKy.getElementsByTagName("co")[0]) : null;
      }

      if (targetSection) {
        let exactVal: number | null = null;
        let subSum = 0;
        let subCount = 0;

        for (let i = 0; i < targetSection.children.length; i++) {
          const c = targetSection.children[i];
          const tag = c.localName.toLowerCase();
          if (tag === `ct${tkPrefix}` || tag === `ct_${tkPrefix}`) {
            exactVal = parseNum(c.textContent);
          } else if (tag.startsWith(`ct${tkPrefix}`) || tag.startsWith(`ct_${tkPrefix}`)) {
            subSum += parseNum(c.textContent);
            subCount++;
          }
        }

        if (exactVal !== null && exactVal !== 0) return exactVal;
        if (subCount > 0 && subSum !== 0) return subSum;
        if (exactVal !== null) return exactVal;
      }
    }

    // B. Cấu trúc phẳng / eTax: <ChiTietCDTK> hoặc attribute maTK
    const allEls = xmlDoc.getElementsByTagName("*");
    let total = 0;
    for (let i = 0; i < allEls.length; i++) {
      const el = allEls[i];
      const tkAttr = el.getAttribute("maTK") || el.getAttribute("MaTK") || el.getAttribute("soHieuTK") ||
                     el.getElementsByTagName("maTK")[0]?.textContent ||
                     el.getElementsByTagName("soHieuTK")[0]?.textContent;
      if (tkAttr && tkAttr.trim().startsWith(tkPrefix)) {
        let val = 0;
        if (field === "duNoDau") {
          val = parseNum(el.getAttribute("duNoDau") || el.getElementsByTagName("duNoDauKy")[0]?.textContent || el.getElementsByTagName("duNoDau")[0]?.textContent);
        } else if (field === "duNoCuoi") {
          val = parseNum(el.getAttribute("duNoCuoi") || el.getElementsByTagName("duNoCuoiKy")[0]?.textContent || el.getElementsByTagName("duNoCuoi")[0]?.textContent);
        } else if (field === "psNo") {
          val = parseNum(el.getAttribute("psNo") || el.getElementsByTagName("soPhatSinhNo")[0]?.textContent || el.getElementsByTagName("psNo")[0]?.textContent);
        } else if (field === "psCo") {
          val = parseNum(el.getAttribute("psCo") || el.getElementsByTagName("soPhatSinhCo")[0]?.textContent || el.getElementsByTagName("psCo")[0]?.textContent);
        }
        total += val;
      }
    }
    return total;
  };

  // 1. Doanh thu thuần (Mã 10 KQKD hoặc TK 511 hoặc Mã 01 KQKD)
  let dtt_10 = getKqkdValue("10");
  if (dtt_10 === 0) dtt_10 = getCdtkAccountAmount("511", "psCo");
  if (dtt_10 === 0) dtt_10 = getKqkdValue("01");

  // 2. Giá vốn hàng bán (Mã 11 KQKD hoặc TK 632)
  let gvhb_11 = getKqkdValue("11");
  if (gvhb_11 === 0) gvhb_11 = getCdtkAccountAmount("632", "psNo");

  // 3. Chi phí bán hàng (Mã 25 KQKD hoặc TK 641/6421)
  let cfbh_25 = getKqkdValue("25");
  if (cfbh_25 === 0) {
    cfbh_25 = getCdtkAccountAmount("641", "psNo") || getCdtkAccountAmount("6421", "psNo");
  }

  // 4. Chi phí quản lý (Mã 24 ở TT133, Mã 26 ở TT200 hoặc TK 642)
  let cfql_26 = thongTu === "TT133" ? getKqkdValue("24") : getKqkdValue("26");
  if (cfql_26 === 0) cfql_26 = getKqkdValue("24") || getKqkdValue("26");
  if (cfql_26 === 0) {
    cfql_26 = getCdtkAccountAmount("642", "psNo") || getCdtkAccountAmount("6422", "psNo");
  }

  // 5. Chi phí lãi vay (Mã 23 KQKD hoặc TK 635)
  let laiVay_23 = getKqkdValue("23");
  if (laiVay_23 === 0) laiVay_23 = getCdtkAccountAmount("635", "psNo");

  // 6. Lợi nhuận
  let ln_thuan_30 = getKqkdValue("30");
  if (ln_thuan_30 === 0 && dtt_10 > 0) {
    ln_thuan_30 = dtt_10 - gvhb_11 - cfbh_25 - cfql_26;
  }
  let ln_truoc_thue_50 = getKqkdValue("50");
  if (ln_truoc_thue_50 === 0) ln_truoc_thue_50 = ln_thuan_30;

  // 7. Hàng tồn kho & Dở dang (TK 154, 155, 157, 156)
  // Ưu tiên đọc từ Bảng Cân đối tài khoản
  let cfsxdd_154_dauKy = getCdtkAccountAmount("154", "duNoDau");
  let cfsxdd_154_cuoiKy = getCdtkAccountAmount("154", "duNoCuoi");
  const thanhPham_155_dauKy = getCdtkAccountAmount("155", "duNoDau");
  const thanhPham_155_cuoiKy = getCdtkAccountAmount("155", "duNoCuoi");
  const hangGuiBan_157_dauKy = getCdtkAccountAmount("157", "duNoDau");
  const hangGuiBan_157_cuoiKy = getCdtkAccountAmount("157", "duNoCuoi");
  const hangHoa_156_dauKy = getCdtkAccountAmount("156", "duNoDau");
  const hangHoa_156_cuoiKy = getCdtkAccountAmount("156", "duNoCuoi");

  // Hàng tồn kho trên Bảng CĐKT (Mã 140 hoặc 141)
  let tonKho_141 = getCdktValue("141", "cuoiNam") || getCdktValue("140", "cuoiNam");
  const tonKho_141_dau = getCdktValue("141", "dauNam") || getCdktValue("140", "dauNam");

  // Fallback: nếu TK 154 chưa có trong CDTK nhưng CĐKT có số dư tồn kho (như DN xây dựng, gia công)
  if (cfsxdd_154_dauKy === 0 && thanhPham_155_dauKy === 0 && hangHoa_156_dauKy === 0 && tonKho_141_dau > 0) {
    cfsxdd_154_dauKy = tonKho_141_dau;
  }
  if (cfsxdd_154_cuoiKy === 0 && thanhPham_155_cuoiKy === 0 && hangHoa_156_cuoiKy === 0 && tonKho_141 > 0) {
    cfsxdd_154_cuoiKy = tonKho_141;
  }

  // 8. Chi phí sản xuất chi tiết (TK 621, 622, 627, 623 hoặc phát sinh TK 154)
  const psNo_154 = getCdtkAccountAmount("154", "psNo");
  let cf_nvl_621 = getCdtkAccountAmount("621", "psNo");
  let cf_nc_622 = getCdtkAccountAmount("622", "psNo");
  let cf_sxc_627 = getCdtkAccountAmount("627", "psNo");
  const cf_mtc_623 = getCdtkAccountAmount("623", "psNo");

  // Nếu TT133 không có 621, 622, 627 mà có phát sinh Nợ TK 154 (Chi phí sản xuất trong kỳ):
  if (thongTu === "TT133" && cf_nvl_621 === 0 && cf_nc_622 === 0 && cf_sxc_627 === 0 && psNo_154 > 0) {
    // Phân bổ ước tính sơ bộ theo tỷ lệ sản xuất công nghiệp chuẩn: NVL ~ 65%, Nhân công ~ 20%, Chung ~ 15%
    cf_nvl_621 = Math.round(psNo_154 * 0.65);
    cf_nc_622 = Math.round(psNo_154 * 0.20);
    cf_sxc_627 = Math.round(psNo_154 * 0.15);
  }

  // 9. Khấu hao (TK 214 phát sinh Có hoặc CĐKT mã 152 lũy kế)
  let khauHao_214 = getCdtkAccountAmount("214", "psCo");
  if (khauHao_214 === 0) {
    const khCuoi = Math.abs(getCdktValue("152", "cuoiNam"));
    const khDau = Math.abs(getCdktValue("152", "dauNam"));
    if (khCuoi >= khDau && khCuoi > 0) {
      khauHao_214 = khCuoi - khDau;
    }
  }

  // 10. Thuế & Lệ phí (TK 333 phát sinh Có hoặc Thuế TNDN mã 51 KQKD)
  let thue_333 = getCdtkAccountAmount("333", "psCo");
  if (thue_333 === 0) {
    thue_333 = getKqkdValue("51");
  }

  // 11. Tiền lương (TK 334 phát sinh Có)
  const tienLuong_334 = getCdtkAccountAmount("334", "psCo");

  // 12. Bảo hiểm (TK 338 phát sinh Có)
  const bhxh_338 = getCdtkAccountAmount("338", "psCo");

  // 13. Nợ vay
  const vayNganHan_320 = getCdktValue("311", "cuoiNam") || getCdktValue("320", "cuoiNam");
  const vayDaiHan_338 = getCdktValue("320", "cuoiNam") || getCdktValue("338", "cuoiNam");

  return {
    mst: mst || "0101437269",
    tenDoanhNghiep: tenDoanhNghiep || fileName.replace(/\.[^/.]+$/, ""),
    diaChi: diaChi || "Tỉnh Hưng Yên",
    nam,
    thongTu,
    dtt_10,
    gvhb_11,
    cfbh_25,
    cfql_26,
    laiVay_23,
    ln_thuan_30,
    ln_truoc_thue_50,
    tonKho_141,
    cfsxdd_154_dauKy,
    cfsxdd_154_cuoiKy,
    thanhPham_155_dauKy,
    thanhPham_155_cuoiKy,
    hangGuiBan_157_dauKy,
    hangGuiBan_157_cuoiKy,
    hangHoa_156_dauKy,
    hangHoa_156_cuoiKy,
    cf_nvl_621,
    cf_nc_622,
    cf_sxc_627,
    cf_mtc_623,
    khauHao_214,
    thue_333,
    tienLuong_334,
    bhxh_338,
    vayNganHan_320,
    vayDaiHan_338,
    sourceFileName: fileName,
    sourceType: "xml"
  };
}

// 2. Phân tích nội dung PDF BCTC (Trích xuất văn bản & regex)
export function parseBctcPdfText(pdfText: string, fileName: string): RawBctcData {
  const clean = pdfText.replace(/\r\n/g, "\n");

  const extractRegex = (regex: RegExp): string => {
    const match = clean.match(regex);
    return match ? match[1]?.trim() || "" : "";
  };

  const extractNumber = (patterns: RegExp[]): number => {
    for (const pat of patterns) {
      const match = clean.match(pat);
      if (match && match[1]) {
        const raw = match[1].replace(/[\.,\s]/g, "").replace(/\((.*?)\)/, "-$1");
        const num = Number(raw);
        if (!isNaN(num) && num !== 0) return num;
      }
    }
    return 0;
  };

  // Trích xuất thông tin cơ bản
  const mst =
    extractRegex(/Mã\s*số\s*thuế\s*[:\.]?\s*([0-9\s\-]+)/i) ||
    extractRegex(/MST\s*[:\.]?\s*([0-9\s\-]+)/i);

  const tenDoanhNghiep =
    extractRegex(/Tên\s*doanh\s*nghiệp\s*[:\.]?\s*([^\n\r]+)/i) ||
    extractRegex(/Tên\s*công\s*ty\s*[:\.]?\s*([^\n\r]+)/i) ||
    fileName.replace(/\.[^/.]+$/, "");

  const diaChi = extractRegex(/Địa\s*chỉ\s*[:\.]?\s*([^\n\r]+)/i);

  // Nhận diện mã số kết quả hoạt động kinh doanh (Mã số 01, 10, 11, 20, 22, 23, 25, 26, 30)
  const dtt_10 = extractNumber([
    /Doanh\s*thu\s*thuần.*?10\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*10.*?([0-9\.,\s]{4,})/i,
    /Doanh\s*thu\s*thuần\s*về\s*bán\s*hàng.*?([0-9\.,\s]{6,})/i
  ]);

  const gvhb_11 = extractNumber([
    /Giá\s*vốn\s*hàng\s*bán.*?11\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*11.*?([0-9\.,\s]{4,})/i
  ]);

  const cfbh_25 = extractNumber([
    /Chi\s*phí\s*bán\s*hàng.*?25\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*25.*?([0-9\.,\s]{4,})/i
  ]);

  const cfql_26 = extractNumber([
    /Chi\s*phí\s*quản\s*lý.*?26\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*26.*?([0-9\.,\s]{4,})/i
  ]);

  const laiVay_23 = extractNumber([
    /Chi\s*phí\s*lãi\s*vay.*?23\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*23.*?([0-9\.,\s]{4,})/i
  ]);

  const ln_thuan_30 = extractNumber([
    /Lợi\s*nhuận\s*thuần.*?30\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*30.*?([0-9\.,\s]{4,})/i
  ]) || (dtt_10 - gvhb_11 - cfbh_25 - cfql_26);

  const ln_truoc_thue_50 = extractNumber([
    /Tổng\s*lợi\s*nhuận\s*kế\s*toán\s*trước\s*thuế.*?50\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*50.*?([0-9\.,\s]{4,})/i
  ]) || ln_thuan_30;

  const tonKho_141 = extractNumber([
    /Hàng\s*tồn\s*kho.*?141\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*141.*?([0-9\.,\s]{4,})/i
  ]);

  const khauHao_214 = extractNumber([
    /Hao\s*mòn\s*lũy\s*kế.*?214\s+[\(]?([0-9\.,\s]+)[\)]?/i,
    /Khấu\s*hao.*?([0-9\.,\s]{6,})/i
  ]);

  const vayNganHan_320 = extractNumber([
    /Vay\s*và\s*nợ\s*thuê\s*tài\s*chính\s*ngắn\s*hạn.*?320\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*320.*?([0-9\.,\s]{4,})/i
  ]);

  const vayDaiHan_338 = extractNumber([
    /Vay\s*và\s*nợ\s*thuê\s*tài\s*chính\s*dài\s*hạn.*?338\s+([0-9\.,\s]+)/i,
    /Mã\s*số\s*338.*?([0-9\.,\s]{4,})/i
  ]);

  return {
    mst: mst || "0100109106",
    tenDoanhNghiep: tenDoanhNghiep || "Công ty Cổ phần Mẫu",
    diaChi: diaChi || "Thành phố Hưng Yên, Tỉnh Hưng Yên",
    nam: 2025,
    thongTu: clean.includes("133/2016") ? "TT133" : "TT200",
    dtt_10,
    gvhb_11,
    cfbh_25,
    cfql_26,
    laiVay_23,
    ln_thuan_30,
    ln_truoc_thue_50,
    tonKho_141,
    cfsxdd_154_dauKy: 0,
    cfsxdd_154_cuoiKy: 0,
    thanhPham_155_dauKy: 0,
    thanhPham_155_cuoiKy: 0,
    hangGuiBan_157_dauKy: 0,
    hangGuiBan_157_cuoiKy: 0,
    hangHoa_156_dauKy: 0,
    hangHoa_156_cuoiKy: 0,
    cf_nvl_621: 0,
    cf_nc_622: 0,
    cf_sxc_627: 0,
    cf_mtc_623: 0,
    khauHao_214,
    thue_333: 0,
    tienLuong_334: 0,
    bhxh_338: 0,
    vayNganHan_320,
    vayDaiHan_338,
    sourceFileName: fileName,
    sourceType: "pdf"
  };
}

export interface IndustrySectorProfile {
  id: string;
  name: string;
  defaultIoCode: string;
  defaultC5Code: string;
  description: string;
  keywords: string[];
  nvlShareInIc: number; // Tỷ trọng NVL trong chi phí trung gian (IC)
  materialRatios: Array<{
    maIO: string;
    moTa: string;
    ratio: number;
    tyLeNK?: number;
  }>;
  serviceRatios: Array<{
    maIO: string;
    moTa: string;
    ratio: number;
    tyLeNK?: number;
  }>;
}

export const INDUSTRY_PROFILES: IndustrySectorProfile[] = [
  {
    id: "xay_dung",
    name: "Xây dựng công trình nhà các loại và kỹ thuật",
    defaultIoCode: "105",
    defaultC5Code: "41000",
    description: "Doanh nghiệp xây dựng dân dụng, công nghiệp, cầu đường, xây lắp, hạ tầng kỹ thuật",
    keywords: [
      "xây dựng", "xây lắp", "thi công", "hạ tầng", "cầu đường", "giao thông",
      "kiến trúc", "đầu tư và xây dựng", "bê tông", "san lấp", "công trình", "kết cấu xây dựng"
    ],
    nvlShareInIc: 0.72,
    materialRatios: [
      { maIO: "075", moTa: "Thép xây dựng, sắt cây, thép hình, kết cấu kim loại", ratio: 0.28 },
      { maIO: "073", moTa: "Xi măng, vôi, thạch cao", ratio: 0.20 },
      { maIO: "074", moTa: "Bê tông thương phẩm, vữa và cấu kiện đúc sẵn", ratio: 0.18 },
      { maIO: "060", moTa: "Cát vàng, đá dăm, sỏi, đất san lấp móng", ratio: 0.14 },
      { maIO: "072", moTa: "Gạch xây, gạch lát, ngói, gốm sứ hoàn thiện", ratio: 0.08 },
      { maIO: "070", moTa: "Sơn tường, bột bả, hóa chất chống thấm", ratio: 0.06 },
      { maIO: "068", moTa: "Xăng dầu diesel vận hành máy xúc, lu, ủi thi công", ratio: 0.06 }
    ],
    serviceRatios: [
      { maIO: "126", moTa: "Cước vận chuyển đường bộ (vận chuyển vật tư, phế thải công trình)", ratio: 0.30 },
      { maIO: "154", moTa: "Thuê máy móc thiết bị thi công, giàn giáo, cốp pha", ratio: 0.25 },
      { maIO: "160", moTa: "Dịch vụ khảo sát thiết kế, tư vấn giám sát công trình", ratio: 0.15 },
      { maIO: "115", moTa: "Tiền điện sinh hoạt & thi công công trường", ratio: 0.10 },
      { maIO: "149", moTa: "Phí dịch vụ ngân hàng, bảo lãnh dự thầu, bảo lãnh thực hiện", ratio: 0.08 },
      { maIO: "146", moTa: "Dịch vụ bưu chính viễn thông, mạng liên lạc", ratio: 0.05 },
      { maIO: "118", moTa: "Nước sạch phục vụ công trường & sinh hoạt", ratio: 0.04 },
      { maIO: "179", moTa: "Sửa chữa, bảo dưỡng phương tiện thi công cơ giới", ratio: 0.03 }
    ]
  },
  {
    id: "thuc_pham_nong_san",
    name: "Chế biến nông sản, thực phẩm và đồ uống",
    defaultIoCode: "024",
    defaultC5Code: "10300",
    description: "Doanh nghiệp sản xuất chế biến nông sản, thủy sản, thực phẩm đóng hộp, bánh kẹo",
    keywords: [
      "thực phẩm", "nông sản", "thủy sản", "hải sản", "chế biến", "bánh kẹo",
      "nước giải khát", "bia", "rượu", "đồ hộp", "thịt", "gạo", "chăn nuôi", "thức ăn"
    ],
    nvlShareInIc: 0.75,
    materialRatios: [
      { maIO: "001", moTa: "Nông sản, rau củ quả, thủy hải sản nguyên liệu thô", ratio: 0.62 },
      { maIO: "055", moTa: "Thùng carton, bao bì giấy đóng gói xuất xưởng", ratio: 0.12 },
      { maIO: "058", moTa: "Màng bọc thực phẩm, túi nilon, khay hộp nhựa", ratio: 0.10 },
      { maIO: "069", moTa: "Phụ gia thực phẩm, gia vị, hương liệu an toàn", ratio: 0.06 },
      { maIO: "068", moTa: "Nhiên liệu khí đốt lò hơi, dầu chạy máy phát", ratio: 0.06 },
      { maIO: "073", moTa: "Muối khoáng, chất bảo quản", ratio: 0.04 }
    ],
    serviceRatios: [
      { maIO: "115", moTa: "Tiền điện kho lạnh bảo quản và dây chuyền chế biến", ratio: 0.38 },
      { maIO: "126", moTa: "Cước vận tải phân phối đường bộ (xe lạnh, xe tải)", ratio: 0.24 },
      { maIO: "118", moTa: "Nước sạch phục vụ chế biến & xử lý nước thải", ratio: 0.16 },
      { maIO: "179", moTa: "Bảo trì, sửa chữa máy móc đóng gói, dây chuyền", ratio: 0.10 },
      { maIO: "154", moTa: "Thuê kho lạnh, nhà xưởng chế biến", ratio: 0.06 },
      { maIO: "146", moTa: "Dịch vụ viễn thông, mạng phần mềm", ratio: 0.03 },
      { maIO: "149", moTa: "Phí dịch vụ ngân hàng, thanh toán", ratio: 0.03 }
    ]
  },
  {
    id: "co_khi_che_tao",
    name: "Sản xuất sản phẩm từ kim loại đúc sẵn và cơ khí",
    defaultIoCode: "082",
    defaultC5Code: "25900",
    description: "Doanh nghiệp gia công cơ khí, tiện phay mài, kết cấu thép, khuôn mẫu, nhôm kính",
    keywords: [
      "cơ khí", "chế tạo", "kim loại", "khuôn mẫu", "đúc", "tiện",
      "kết cấu thép", "nhôm kính", "inox", "gia công kim loại", "máy móc"
    ],
    nvlShareInIc: 0.70,
    materialRatios: [
      { maIO: "075", moTa: "Thép tấm, thép cuộn, phôi nhôm, kim loại cơ bản", ratio: 0.60 },
      { maIO: "082", moTa: "Que hàn, dây hàn, bulong, đai ốc, phụ kiện lắp ráp", ratio: 0.15 },
      { maIO: "070", moTa: "Sơn chống rỉ sét, hóa chất tẩy mạ kẽm bề mặt", ratio: 0.10 },
      { maIO: "068", moTa: "Dầu tưới nguội cắt gọt CNC, dầu thủy lực bôi trơn", ratio: 0.08 },
      { maIO: "055", moTa: "Vật liệu đai kẹp pallet, màng quấn bảo vệ", ratio: 0.07 }
    ],
    serviceRatios: [
      { maIO: "115", moTa: "Tiền điện sản xuất 3 pha (lò nung, máy hàn, máy cắt CNC)", ratio: 0.40 },
      { maIO: "179", moTa: "Dịch vụ gia công ngoài (tiện mài, nhiệt luyện, mạ kẽm)", ratio: 0.25 },
      { maIO: "126", moTa: "Cước xe cẩu và vận tải cấu kiện sắt thép cơ khí", ratio: 0.15 },
      { maIO: "154", moTa: "Thuê mặt bằng nhà xưởng cơ khí", ratio: 0.10 },
      { maIO: "146", moTa: "Dịch vụ viễn thông, bản quyền phần mềm CAD/CAM", ratio: 0.05 },
      { maIO: "149", moTa: "Phí ngân hàng, thanh toán thương mại", ratio: 0.05 }
    ]
  },
  {
    id: "det_may_giay_da",
    name: "Sản xuất trang phục và sản phẩm dệt may da giày",
    defaultIoCode: "045",
    defaultC5Code: "14100",
    description: "Doanh nghiệp dệt may, may mặc thời trang xuất khẩu, da giày",
    keywords: [
      "may", "dệt", "thời trang", "giày", "da", "sợi", "quần áo", "may mặc", "vải"
    ],
    nvlShareInIc: 0.68,
    materialRatios: [
      { maIO: "044", moTa: "Vải cuộn các loại (cotton, polyester, vải dệt kim)", ratio: 0.58 },
      { maIO: "045", moTa: "Phụ liệu may mặc (chỉ may, cúc, khóa kéo, nhãn mác, chun)", ratio: 0.18 },
      { maIO: "055", moTa: "Thùng carton xuất khẩu, bìa cứng định hình", ratio: 0.10 },
      { maIO: "058", moTa: "Túi nilon bọc áo, gói hút ẩm bảo quản", ratio: 0.08 },
      { maIO: "069", moTa: "Thuốc nhuộm, hóa chất xử lý mềm vải", ratio: 0.06 }
    ],
    serviceRatios: [
      { maIO: "115", moTa: "Tiền điện xưởng may (hệ thống máy may, chiếu sáng)", ratio: 0.35 },
      { maIO: "140", moTa: "Dịch vụ in chuyển nhiệt, thêu vi tính, giặt là ngoài", ratio: 0.25 },
      { maIO: "126", moTa: "Cước vận chuyển container, xe tải giao hàng", ratio: 0.18 },
      { maIO: "154", moTa: "Thuê nhà xưởng sản xuất may mặc", ratio: 0.12 },
      { maIO: "146", moTa: "Cước viễn thông, chứng chỉ CO xuất xứ", ratio: 0.05 },
      { maIO: "149", moTa: "Phí ngân hàng mở thư tín dụng L/C", ratio: 0.05 }
    ]
  },
  {
    id: "thuong_mai_ban_buon_le",
    name: "Bán buôn và bán lẻ hàng hóa, thương mại phân phối",
    defaultIoCode: "124",
    defaultC5Code: "46900",
    description: "Doanh nghiệp thương mại, xuất nhập khẩu, đại lý phân phối hàng hóa tiêu dùng, vật tư",
    keywords: [
      "thương mại", "xuất nhập khẩu", "phân phối", "siêu thị", "bán buôn",
      "bán lẻ", "cửa hàng", "kinh doanh tổng hợp", "đại lý"
    ],
    nvlShareInIc: 0.35,
    materialRatios: [
      { maIO: "055", moTa: "Thùng carton, hộp giấy bao bì đóng gói hàng hóa", ratio: 0.45 },
      { maIO: "058", moTa: "Túi đựng hàng, bao nilon, màng quấn PE, băng dính", ratio: 0.35 },
      { maIO: "068", moTa: "Xăng dầu xe phục vụ đi thị trường, chào hàng phân phối", ratio: 0.20 }
    ],
    serviceRatios: [
      { maIO: "154", moTa: "Thuê mặt bằng showroom, cửa hàng, kho trung chuyển", ratio: 0.35 },
      { maIO: "126", moTa: "Cước vận tải giao nhận hàng hóa (shipper, xe tải)", ratio: 0.25 },
      { maIO: "158", moTa: "Chi phí quảng cáo tiếp thị, bảng hiệu, khuyến mãi", ratio: 0.15 },
      { maIO: "115", moTa: "Tiền điện showroom, điều hòa nhiệt độ, chiếu sáng", ratio: 0.12 },
      { maIO: "149", moTa: "Phí máy quẹt thẻ POS, thanh toán ngân hàng điện tử", ratio: 0.08 },
      { maIO: "146", moTa: "Cước viễn thông, Internet và phần mềm quản lý bán hàng", ratio: 0.05 }
    ]
  },
  {
    id: "van_tai_logistics",
    name: "Vận tải hàng hóa đường bộ và dịch vụ logistics kho bãi",
    defaultIoCode: "126",
    defaultC5Code: "49330",
    description: "Doanh nghiệp vận tải xe tải, xe container, dịch vụ giao nhận, kho bãi bốc xếp",
    keywords: [
      "vận tải", "logistics", "kho bãi", "xe tải", "xe khách",
      "chuyển phát", "giao nhận", "vận chuyển"
    ],
    nvlShareInIc: 0.55,
    materialRatios: [
      { maIO: "068", moTa: "Dầu diesel, xăng xe tải đầu kéo đường dài", ratio: 0.75 },
      { maIO: "057", moTa: "Lốp xe, săm xe tải cơ giới thay thế định kỳ", ratio: 0.15 },
      { maIO: "068", moTa: "Dầu nhớt bôi trơn động cơ, mỡ bò chịu nhiệt", ratio: 0.10 }
    ],
    serviceRatios: [
      { maIO: "126", moTa: "Phí đường bộ, vé cầu đường BOT, vé bến bãi đỗ xe", ratio: 0.40 },
      { maIO: "179", moTa: "Dịch vụ garage sửa chữa, thay thế phụ tùng xe tải", ratio: 0.25 },
      { maIO: "147", moTa: "Phí bảo hiểm trách nhiệm dân sự xe và bảo hiểm hàng hóa", ratio: 0.15 },
      { maIO: "154", moTa: "Thuê bãi đỗ xe tải, kho trung chuyển hàng hóa", ratio: 0.10 },
      { maIO: "146", moTa: "Cước thiết bị định vị giám sát hành trình GPS, 4G", ratio: 0.06 },
      { maIO: "149", moTa: "Phí dịch vụ ngân hàng, thu hộ tiền hàng COD", ratio: 0.04 }
    ]
  },
  {
    id: "cong_nghe_thong_tin",
    name: "Lập trình máy tính và dịch vụ công nghệ thông tin",
    defaultIoCode: "145",
    defaultC5Code: "62010",
    description: "Doanh nghiệp phần mềm, dịch vụ CNTT, giải pháp số, website, máy chủ đám mây",
    keywords: [
      "công nghệ", "phần mềm", "tin học", "it", "số hóa",
      "truyền thông", "mạng", "phần cứng"
    ],
    nvlShareInIc: 0.15,
    materialRatios: [
      { maIO: "055", moTa: "Văn phòng phẩm, giấy in tài liệu hợp đồng", ratio: 0.50 },
      { maIO: "085", moTa: "Thiết bị linh kiện ngoại vi, phụ kiện mạng", ratio: 0.50 }
    ],
    serviceRatios: [
      { maIO: "145", moTa: "Thuê máy chủ Cloud (AWS, GCP, Viettel), bản quyền phần mềm", ratio: 0.45 },
      { maIO: "154", moTa: "Thuê văn phòng làm việc, không gian sáng tạo", ratio: 0.25 },
      { maIO: "146", moTa: "Đường truyền Internet cáp quang tốc độ cao", ratio: 0.15 },
      { maIO: "115", moTa: "Tiền điện chạy điều hòa và máy chủ văn phòng", ratio: 0.08 },
      { maIO: "149", moTa: "Phí cổng thanh toán trực tuyến, tài khoản ngân hàng", ratio: 0.07 }
    ]
  },
  {
    id: "san_xuat_chung",
    name: "Sản xuất công nghiệp tổng hợp khác",
    defaultIoCode: "099",
    defaultC5Code: "32900",
    description: "Doanh nghiệp sản xuất chế biến khác không thuộc các nhóm đặc thù trên",
    keywords: [],
    nvlShareInIc: 0.65,
    materialRatios: [
      { maIO: "075", moTa: "Nguyên vật liệu chính", ratio: 0.50 },
      { maIO: "055", moTa: "Bao bì, đóng gói sản phẩm", ratio: 0.25 },
      { maIO: "068", moTa: "Xăng dầu và nhiên liệu vận hành", ratio: 0.15 },
      { maIO: "070", moTa: "Vật tư phụ, hóa chất", ratio: 0.10 }
    ],
    serviceRatios: [
      { maIO: "115", moTa: "Tiền điện sản xuất và văn phòng", ratio: 0.35 },
      { maIO: "126", moTa: "Cước vận tải hàng hóa", ratio: 0.25 },
      { maIO: "154", moTa: "Thuê mặt bằng, xưởng sản xuất", ratio: 0.15 },
      { maIO: "179", moTa: "Dịch vụ sửa chữa, bảo dưỡng máy móc", ratio: 0.15 },
      { maIO: "146", moTa: "Dịch vụ viễn thông, liên lạc", ratio: 0.05 },
      { maIO: "149", moTa: "Phí dịch vụ ngân hàng", ratio: 0.05 }
    ]
  }
];

// Hàm nhận diện ngành nghề thông minh dựa trên tên công ty hoặc mã IO tùy chọn
export function detectIndustrySector(tenDoanhNghiep: string, customIo?: string): IndustrySectorProfile {
  if (customIo) {
    const matched = INDUSTRY_PROFILES.find(p => p.defaultIoCode === customIo);
    if (matched) return matched;
  }
  const clean = (tenDoanhNghiep || "").toLowerCase();
  for (const prof of INDUSTRY_PROFILES) {
    if (prof.id === "san_xuat_chung") continue;
    for (const kw of prof.keywords) {
      if (clean.includes(kw)) {
        return prof;
      }
    }
  }
  return INDUSTRY_PROFILES.find(p => p.id === "san_xuat_chung") || INDUSTRY_PROFILES[0];
}

// 3. Chuyển đổi dữ liệu BCTC vào MẪU PHIẾU 01/IO-DN
// BÓC TÁCH THÔNG MINH TỪ TỔNG VÀO CÁC MÃ IO THEO ĐÚNG NGÀNH NGHỀ
export function convertBctcToPhieu01IODN(
  bctc: RawBctcData,
  customMainIo?: string,
  customSectorId?: string
): Phieu01IODNForm {
  // 1. Xác định ngành nghề thực tế của doanh nghiệp
  let sector: IndustrySectorProfile;
  if (customSectorId) {
    sector = INDUSTRY_PROFILES.find(p => p.id === customSectorId) || detectIndustrySector(bctc.tenDoanhNghiep, customMainIo);
  } else {
    sector = detectIndustrySector(bctc.tenDoanhNghiep, customMainIo && customMainIo !== "105" ? customMainIo : undefined);
  }

  const actualIoCode = customMainIo && customMainIo !== "105" ? customMainIo : sector.defaultIoCode;
  const dtt = bctc.dtt_10 || 0;
  const gvhb = bctc.gvhb_11 || 0;
  const cfbh = bctc.cfbh_25 || 0;
  const cfql = bctc.cfql_26 || 0;
  const laiVay = bctc.laiVay_23 || 0;
  const khauHao = bctc.khauHao_214 || 0;
  const tienLuong = bctc.tienLuong_334 || 0;
  const bhxh = bctc.bhxh_338 || (tienLuong > 0 ? Math.round(tienLuong * 0.235) : 0);
  const thue = bctc.thue_333 || 0;
  const ln223 = dtt - gvhb - cfbh - cfql;

  // 2. Phân bổ cơ cấu chi phí sản xuất (CFNVL, CFNC, CFSXC, CFMTC)
  let cfnvl = bctc.cf_nvl_621;
  let cfnc = bctc.cf_nc_622;
  let cfsxc = bctc.cf_sxc_627;
  let cfmtc = bctc.cf_mtc_623;

  // Nếu trên BCTC chưa tách riêng 621, 622, 627 (như TT133 tập hợp chung vào TK 154 hoặc chỉ có GVHB):
  // Hệ thống bóc tách thông minh từ Tổng giá vốn theo đúng tỷ trọng đặc thù ngành nghề đã nhận diện
  if ((cfnvl === 0 && cfsxc === 0) || cfnvl === undefined) {
    const tongChiPhi = gvhb > 0 ? gvhb : Math.round(dtt * 0.85);
    const luongSanXuat = tienLuong > 0 ? tienLuong : Math.round(tongChiPhi * 0.12);
    cfnc = luongSanXuat;

    // Chi phí trung gian phân bổ từ giá vốn trừ lương và khấu hao
    const poolTrungGian = Math.max(0, tongChiPhi - cfnc - khauHao);
    cfnvl = Math.round(poolTrungGian * sector.nvlShareInIc);
    const phanConLai = poolTrungGian - cfnvl;

    if (sector.id === "xay_dung") {
      cfmtc = Math.round(phanConLai * 0.35);
      cfsxc = phanConLai - cfmtc;
    } else {
      cfmtc = 0;
      cfsxc = phanConLai;
    }
  }

  // Tồn kho câu 8
  const sxdd_dau = bctc.cfsxdd_154_dauKy || 0;
  const sxdd_cuoi = bctc.cfsxdd_154_cuoiKy || 0;
  const tp_dau = bctc.thanhPham_155_dauKy || 0;
  const tp_cuoi = bctc.thanhPham_155_cuoiKy || 0;
  const gui_dau = bctc.hangGuiBan_157_dauKy || 0;
  const gui_cuoi = bctc.hangGuiBan_157_cuoiKy || 0;
  const tongTonKhoDau = sxdd_dau + tp_dau + gui_dau;
  const tongTonKhoCuoi = sxdd_cuoi + tp_cuoi + gui_cuoi;

  // 3. Bóc tách chi tiết Câu 15: Chi phí Nguyên vật liệu theo từng Mã IO vật tư chuẩn
  const cau15_nvlChiTiet: Array<{ maIO: string; moTa: string; c1_giaTri: number; c2_tyLeNK: number; c3_giaCong: number }> = [];
  let daPhanBoNvl = 0;
  sector.materialRatios.forEach((item, idx) => {
    let val = Math.round(cfnvl * item.ratio);
    if (idx === sector.materialRatios.length - 1) {
      val = Math.max(0, cfnvl - daPhanBoNvl);
    } else {
      daPhanBoNvl += val;
    }
    cau15_nvlChiTiet.push({
      maIO: item.maIO,
      moTa: item.moTa,
      c1_giaTri: val,
      c2_tyLeNK: 0,
      c3_giaCong: 0
    });
  });

  // 4. Bóc tách chi tiết Câu 18: Dịch vụ mua ngoài & chi phí khác theo từng Mã IO dịch vụ
  // Tổng dịch vụ mua ngoài = CFSXC + Chi phí quản lý & bán hàng mua ngoài (khoảng 65% của CFBH+CFQL)
  const tongDichVu = Math.max(0, cfsxc + Math.round((cfbh + cfql) * 0.65));
  const cau18_muaNgoaiChiPhiKhac: Array<{ maIO: string; moTa: string; c1_giaTri: number; c2_tyLeNK: number }> = [];
  let daPhanBoDv = 0;
  sector.serviceRatios.forEach((item, idx) => {
    let val = Math.round(tongDichVu * item.ratio);
    if (idx === sector.serviceRatios.length - 1) {
      val = Math.max(0, tongDichVu - daPhanBoDv);
    } else {
      daPhanBoDv += val;
    }
    cau18_muaNgoaiChiPhiKhac.push({
      maIO: item.maIO,
      moTa: item.moTa,
      c1_giaTri: val,
      c2_tyLeNK: 0
    });
  });

  // 5. Tính IC & GO và tỷ lệ IC/GO
  const ic = cfnvl + cfsxc + cfmtc + Math.round((cfbh + cfql) * 0.65);
  const go = ic + cfnc + khauHao + ln223 + thue;
  const tyLeIC_GO = go > 0 ? ic / go : 0;

  const isSpecialIO = ["115", "116", "129", "130", "153", "142", "164", "176"].includes(actualIoCode);
  let danhGiaIC: "bình thường" | "quá cao (>0.75 hoặc >0.85)" | "quá thấp (<0.3)" = "bình thường";
  if (isSpecialIO) {
    if (tyLeIC_GO > 0.75) danhGiaIC = "quá cao (>0.75 hoặc >0.85)";
    else if (tyLeIC_GO > 0 && tyLeIC_GO < 0.3) danhGiaIC = "quá thấp (<0.3)";
  } else {
    if (tyLeIC_GO > 0.85) danhGiaIC = "quá cao (>0.75 hoặc >0.85)";
    else if (tyLeIC_GO > 0 && tyLeIC_GO < 0.3) danhGiaIC = "quá thấp (<0.3)";
  }

  const form: Phieu01IODNForm = {
    thongTinDinhDanh: {
      tenDoanhNghiep: bctc.tenDoanhNghiep,
      maSoThue: bctc.mst,
      diaChi: bctc.diaChi,
      tinhTP: "33. Tỉnh Hưng Yên",
      soDienThoai: "0221.3862xxx",
      loaiHinhKinhTe: "2", // 2: Doanh nghiệp ngoài nhà nước
      maNganhC5: sector.defaultC5Code,
      maIO: actualIoCode,
      tenNganhChinh: sector.name,
      sectorId: sector.id,
      sectorExplanation: `Tự động nhận diện từ tên doanh nghiệp: "${bctc.tenDoanhNghiep}" -> Thuộc nhóm ${sector.name} (Mã IO: ${actualIoCode}).`
    },
    phan1: {
      cau5_tongDoanhThuThuan: dtt,
      cau6_doanhThuSanPham: [
        {
          stt: 1,
          tenNganh: sector.name,
          maIO: actualIoCode,
          c1_dtt: dtt,
          c2_coGiaCong: 2, // 2: Không
          c3_phiGiaCong: 0
        },
        {
          stt: 2,
          tenNganh: "Sản phẩm phụ / dịch vụ khác",
          maIO: "0",
          c1_dtt: 0,
          c2_coGiaCong: 2,
          c3_phiGiaCong: 0
        }
      ],
      cau7_thongTinKhac: [
        {
          stt: 1,
          tenNganh: sector.name,
          maIO: actualIoCode,
          c1_gvhb: gvhb,
          c2_gvChuyenBan: 0,
          c3_chiHo: 0,
          c4_chiThuong: 0
        }
      ],
      cau8_hangTonKhoDN: {
        c81_sxdd_dauKy: sxdd_dau,
        c81_sxdd_cuoiKy: sxdd_cuoi,
        c82_thanhPham_dauKy: tp_dau,
        c82_thanhPham_cuoiKy: tp_cuoi,
        c83_hangGui_dauKy: gui_dau,
        c83_hangGui_cuoiKy: gui_cuoi,
        tong_dauKy: tongTonKhoDau,
        tong_cuoiKy: tongTonKhoCuoi
      },
      cau9_tonKhoSanPham: [
        {
          stt: 1,
          tenNganh: sector.name,
          maIO: actualIoCode,
          c1_sxdd_dau: sxdd_dau,
          c2_sxdd_cuoi: sxdd_cuoi,
          c3_tp_dau: tp_dau,
          c4_tp_cuoi: tp_cuoi,
          c5_gui_dau: gui_dau,
          c6_gui_cuoi: gui_cuoi
        }
      ]
    },
    phan2: {
      cau10_sxkdChinh: {
        dtt,
        gvhb,
        cfbh,
        cfql,
        ma223_loiNhuan: ln223,
        ma224_laiVay: laiVay,
        cfnvl,
        cfnc,
        cfsxc,
        cfmtc,
        c2_dtt: 0,
        c2_gvhb: 0,
        c2_cfbh: 0,
        c2_cfql: 0,
        c2_loiNhuan: 0,
        c2_laiVay: 0,
        c2_cfnvl: 0,
        c2_cfnc: 0,
        c2_cfsxc: 0,
        c2_cfmtc: 0
      },
      cau11_troCap: [
        { maIO: "202", ten: "Trợ giá sản phẩm", soTien: 0 },
        { maIO: "203", ten: "Trợ cấp vận chuyển", soTien: 0 },
        { maIO: "204", ten: "Trợ cấp xuất nhập khẩu", soTien: 0 },
        { maIO: "205", ten: "Trợ cấp quỹ lương", soTien: 0 },
        { maIO: "206", ten: "Trợ cấp giảm ô nhiễm", soTien: 0 },
        { maIO: "207", ten: "Các loại trợ cấp khác", soTien: 0 }
      ],
      cau12_thueLePhi: [
        { maIO: "208", ten: "Thuế VAT hàng nội địa phát sinh", soTien: thue > 0 ? Math.round(thue * 0.7) : 0, duocHachToan: false },
        { maIO: "210", ten: "Thuế TTĐB nội địa", soTien: 0, duocHachToan: false },
        { maIO: "214", ten: "Thuế bảo vệ môi trường", soTien: 0, duocHachToan: false },
        { maIO: "216", ten: "Thuế sử dụng đất", soTien: thue > 0 ? Math.round(thue * 0.1) : 0, duocHachToan: true },
        { maIO: "217", ten: "Lệ phí liên quan sản phẩm chính", soTien: thue > 0 ? Math.round(thue * 0.2) : 0, duocHachToan: true },
        { maIO: "230", ten: "Thuế khác liên quan sản phẩm chính", soTien: 0, duocHachToan: false }
      ],
      cau13_tienThuong191: 0,
      cau14_thuongMai: {
        c1_dtt: 0,
        c2_muaVe: 0,
        c3_vanTai: 0,
        c4_tonKhoDau: 0,
        c5_tonKhoCuoi: 0,
        c6_phiTM: 0
      }
    },
    phan3: {
      cau15_nvlChiTiet,
      cau16_nhanCong: {
        ma182_tienLuong: tienLuong,
        ma183_bhxh: bhxh > 0 ? Math.round(bhxh * 0.75) : 0,
        ma184_bhyt: bhxh > 0 ? Math.round(bhxh * 0.15) : 0,
        ma185_bhtn: bhxh > 0 ? Math.round(bhxh * 0.05) : 0,
        ma186_bhConNguoi: 0,
        ma187_kpcd: bhxh > 0 ? Math.round(bhxh * 0.05) : 0,
        ma195_chiTrucTiep: 0,
        ma196_chiTraKhac: 0,
        tongCong: tienLuong + bhxh,
        gc_182: 0,
        gc_183: 0,
        gc_184: 0,
        gc_185: 0,
        gc_186: 0,
        gc_187: 0,
        gc_195: 0,
        gc_196: 0,
        gc_tong: 0
      },
      cau17_khauHao225: khauHao,
      cau18_muaNgoaiChiPhiKhac
    },
    phan4: {
      cau194_vayNganHan320: bctc.vayNganHan_320 || 0,
      cau194_vayDaiHan338: bctc.vayDaiHan_338 || 0,
      cau191_coTroCap: false,
      cau192_coBaoLanh: false,
      cau193_coVayNN: false
    },
    tongHopIO: {
      ic,
      go,
      tyLeIC_GO: Number(tyLeIC_GO.toFixed(4)),
      danhGiaIC
    }
  };

  return form;
}

// 3b. HÀM PHÂN TÍCH FILE EXCEL PHIẾU 01/IO-DN CỦA ĐIỀU TRA VIÊN (ĐTV)
export function parsePhieu01Excel(fileData: ArrayBuffer | Uint8Array, fileName: string): { form: Phieu01IODNForm; rawBctc?: RawBctcData } {
  const wb = XLSX.read(fileData, { type: "array", cellDates: true, cellFormula: false });

  // Tìm sheet chứa Phiếu 01
  let targetSheetName = wb.SheetNames[0];
  for (const name of wb.SheetNames) {
    const lower = name.toLowerCase();
    if (lower.includes("phieu") || lower.includes("01") || lower.includes("io") || lower.includes("phieu01")) {
      targetSheetName = name;
      break;
    }
  }

  const ws = wb.Sheets[targetSheetName];
  const rows = (XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any[][]) || [];

  const parseNum = (val: any): number => {
    if (typeof val === "number") return isNaN(val) ? 0 : val;
    if (!val) return 0;
    const s = String(val).replace(/[\r\n\t]/g, "").trim();
    if (s === "-" || s === "" || s === "X" || s === "x" || s === "null" || s === "None") return 0;
    const isNeg = s.startsWith("(") && s.endsWith(")");
    const clean = s.replace(/[\(\)]/g, "").replace(/,/g, "").replace(/\s/g, "");
    const n = parseFloat(clean);
    if (isNaN(n)) return 0;
    return isNeg ? -Math.abs(n) : n;
  };

  const getCell = (r: number, c: number): any => {
    if (r < 0 || r >= rows.length) return "";
    const row = rows[r];
    if (!row || c < 0 || c >= row.length) return "";
    return row[c] ?? "";
  };

  const getRowNums = (r: number): number[] => {
    if (r < 0 || r >= rows.length || !rows[r]) return [];
    return rows[r].map(parseNum).filter(n => n !== 0);
  };

  const findRow = (pattern: RegExp, startRow = 0, maxRow = rows.length): number => {
    for (let r = startRow; r < Math.min(rows.length, maxRow); r++) {
      const row = rows[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        if (pattern.test(String(row[c]))) return r;
      }
    }
    return -1;
  };

  // 1. THÔNG TIN ĐỊNH DANH
  let tenDN = "";
  const rTen = findRow(/Tên\s*doanh\s*nghiệp/i);
  if (rTen >= 0) {
    for (let c = 0; c < rows[rTen].length; c++) {
      const val = String(rows[rTen][c]).trim();
      if (val && !/Tên\s*doanh\s*nghiệp/i.test(val)) {
        tenDN = val;
        break;
      }
    }
  }
  if (!tenDN) tenDN = String(getCell(12, 7) || getCell(13, 7) || getCell(12, 6) || "").trim();
  if (!tenDN) tenDN = "Doanh nghiệp khảo sát IO (File ĐTV)";

  let mst = "";
  const rMst = findRow(/Mã\s*số\s*thuế|MST/i);
  if (rMst >= 0) {
    for (let c = 0; c < rows[rMst].length; c++) {
      const val = String(rows[rMst][c]).trim();
      const match = val.match(/\b\d{10}(-\d{3})?\b/);
      if (match) {
        mst = match[0];
        break;
      }
    }
  }
  if (!mst) {
    const fallbackMst = String(getCell(13, 7) || getCell(14, 7) || getCell(13, 6) || "").trim();
    const match = fallbackMst.match(/\b\d{10}(-\d{3})?\b/);
    if (match) mst = match[0];
  }
  if (!mst) mst = "1001275111";

  let diaChi = "";
  const rDc = findRow(/Địa\s*chỉ/i);
  if (rDc >= 0) {
    for (let c = 0; c < rows[rDc].length; c++) {
      const val = String(rows[rDc][c]).trim();
      if (val && !/Địa\s*chỉ/i.test(val)) {
        diaChi = val;
        break;
      }
    }
  }
  if (!diaChi) diaChi = String(getCell(17, 9) || getCell(18, 9) || "Khu công nghiệp").trim();

  let tinhTP = "";
  const rTinh = findRow(/Tỉnh/i);
  if (rTinh >= 0) {
    for (let c = 0; c < rows[rTinh].length; c++) {
      const val = String(rows[rTinh][c]).trim();
      if (val && !/Tỉnh/i.test(val)) {
        tinhTP = val;
        break;
      }
    }
  }
  if (!tinhTP) tinhTP = String(getCell(15, 9) || getCell(16, 9) || "Hưng Yên").trim();

  let soDienThoai = "";
  const rDt = findRow(/Điện\s*thoại/i);
  if (rDt >= 0) {
    for (let c = 0; c < rows[rDt].length; c++) {
      const val = String(rows[rDt][c]).trim();
      if (val && !/Điện\s*thoại/i.test(val)) {
        soDienThoai = val;
        break;
      }
    }
  }
  if (!soDienThoai) soDienThoai = String(getCell(18, 9) || getCell(19, 9) || "02213888999").trim();

  let loaiHinhKinhTe = "2";
  const rLh = findRow(/Loại\s*hình\s*kinh\s*tế/i);
  if (rLh >= 0) {
    for (let c = 0; c < rows[rLh].length; c++) {
      const val = String(rows[rLh][c]).trim();
      if (/nhà\s*nước/i.test(val)) { loaiHinhKinhTe = "1"; break; }
      if (/nước\s*ngoài|fdi/i.test(val)) { loaiHinhKinhTe = "3"; break; }
    }
  }

  // Ngành sản phẩm & Mã IO
  let tenNganh = "";
  let maC5 = "";
  let maIO = "";
  const rNganh = findRow(/Ngành\s*sản\s*phẩm\s*chính|Mã\s*IO/i);
  if (rNganh >= 0) {
    const row = rows[rNganh];
    for (let c = 0; c < row.length; c++) {
      const val = String(row[c]).trim();
      const ioMatch = val.match(/Mã\s*IO[:\s]*(\d{3})/i) || val.match(/\b(\d{3})\b/);
      if (ioMatch && !maIO) maIO = ioMatch[1];
      const c5Match = val.match(/Mã\s*C5[:\s]*(\d{5})/i) || val.match(/\b(\d{5})\b/);
      if (c5Match && !maC5) maC5 = c5Match[1];
      if (val.length > 5 && !/Mã\s*IO|Mã\s*C5/i.test(val) && !tenNganh) {
        tenNganh = val;
      }
    }
  }
  if (!tenNganh) tenNganh = String(getCell(23, 3) || getCell(24, 3) || "").trim();
  if (!maC5) maC5 = String(getCell(23, 21) || getCell(24, 21) || "").trim();
  if (!maIO) maIO = String(getCell(23, 26) || getCell(24, 26) || "").trim();

  const detected = detectIndustrySector(tenDN + " " + tenNganh);
  if (!maIO) maIO = detected.defaultIoCode;
  if (!maC5) maC5 = detected.defaultC5Code;
  if (!tenNganh) tenNganh = detected.name;

  // CÂU 5: TỔNG DOANH THU THUẦN
  let cau5_tongDoanhThuThuan = 0;
  const rC5 = findRow(/5\.\s*Tổng\s*doanh\s*thu|Tổng\s*doanh\s*thu\s*thuần/i);
  if (rC5 >= 0) {
    const nums = getRowNums(rC5);
    if (nums.length > 0) cau5_tongDoanhThuThuan = nums[nums.length - 1];
  }
  if (cau5_tongDoanhThuThuan === 0) {
    cau5_tongDoanhThuThuan = parseNum(getCell(25, 25)) || parseNum(getCell(25, 26)) || parseNum(getCell(26, 26)) || parseNum(getCell(25, 21));
  }

  // CÂU 6: DOANH THU SẢN PHẨM
  const cau6_doanhThuSanPham: Array<{
    stt: number;
    tenNganh: string;
    maIO: string;
    c1_dtt: number;
    c2_coGiaCong: number;
    c3_phiGiaCong: number;
  }> = [];

  const rC6 = findRow(/6\.\s*Doanh\s*thu\s*thuần\s*chia\s*theo/i);
  if (rC6 >= 0) {
    for (let r = rC6 + 1; r < Math.min(rows.length, rC6 + 15); r++) {
      const row = rows[r];
      if (!row) continue;
      const rowStr = row.map(String).join(" ");
      if (/7\.\s*Thông\s*tin\s*khác/i.test(rowStr)) break;
      if (/TỔNG\s*CỘNG/i.test(rowStr)) continue;

      let rowIo = "";
      for (const cell of row) {
        const s = String(cell).trim();
        if (/^\d{3}$/.test(s)) {
          rowIo = s;
          break;
        }
      }
      const nums = getRowNums(r);
      if (nums.length >= 1) {
        const dtt = nums[0];
        let coGiaCong = 2;
        let phiGiaCong = 0;
        if (/có|1/i.test(rowStr) && !/không/i.test(rowStr)) coGiaCong = 1;
        if (nums.length >= 2) phiGiaCong = nums[nums.length - 1];
        if (phiGiaCong === dtt && nums.length < 3) phiGiaCong = 0;

        cau6_doanhThuSanPham.push({
          stt: cau6_doanhThuSanPham.length + 1,
          tenNganh: tenNganh || `Sản phẩm ${rowIo || maIO}`,
          maIO: rowIo || maIO,
          c1_dtt: dtt,
          c2_coGiaCong: coGiaCong,
          c3_phiGiaCong: phiGiaCong
        });
      }
    }
  }

  if (cau6_doanhThuSanPham.length === 0) {
    const rev31 = parseNum(getCell(30, 20)) || parseNum(getCell(30, 21)) || parseNum(getCell(31, 20)) || parseNum(getCell(31, 21));
    const dttVal = rev31 || cau5_tongDoanhThuThuan;
    cau6_doanhThuSanPham.push({
      stt: 1,
      tenNganh,
      maIO,
      c1_dtt: dttVal,
      c2_coGiaCong: 1, // ĐTV thường chọn có gia công
      c3_phiGiaCong: 0
    });
  }
  if (cau5_tongDoanhThuThuan === 0 && cau6_doanhThuSanPham.length > 0) {
    cau5_tongDoanhThuThuan = cau6_doanhThuSanPham.reduce((s, r) => s + r.c1_dtt, 0);
  }

  // CÂU 7: THÔNG TIN KHÁC
  const cau7_thongTinKhac: Array<{
    stt: number;
    tenNganh: string;
    maIO: string;
    c1_gvhb: number;
    c2_gvChuyenBan: number;
    c3_chiHo: number;
    c4_chiThuong: number;
  }> = [];

  const rC7 = findRow(/7\.\s*Thông\s*tin\s*khác/i);
  if (rC7 >= 0) {
    for (let r = rC7 + 1; r < Math.min(rows.length, rC7 + 12); r++) {
      const row = rows[r];
      if (!row) continue;
      const rowStr = row.map(String).join(" ");
      if (/8\.\s*Thông\s*tin/i.test(rowStr)) break;
      if (/TỔNG/i.test(rowStr)) continue;

      let rowIo = "";
      for (const cell of row) {
        const s = String(cell).trim();
        if (/^\d{3}$/.test(s)) {
          rowIo = s;
          break;
        }
      }
      const nums = getRowNums(r);
      if (nums.length >= 1) {
        cau7_thongTinKhac.push({
          stt: cau7_thongTinKhac.length + 1,
          tenNganh: tenNganh || `Sản phẩm ${rowIo || maIO}`,
          maIO: rowIo || maIO,
          c1_gvhb: nums[0] || 0,
          c2_gvChuyenBan: nums[1] || 0,
          c3_chiHo: nums[2] || 0,
          c4_chiThuong: nums[3] || 0
        });
      }
    }
  }
  if (cau7_thongTinKhac.length === 0) {
    cau7_thongTinKhac.push({
      stt: 1,
      tenNganh,
      maIO,
      c1_gvhb: Math.round(cau5_tongDoanhThuThuan * 0.78),
      c2_gvChuyenBan: 0,
      c3_chiHo: 0,
      c4_chiThuong: 0
    });
  }

  // CÂU 8: HÀNG TỒN KHO DN
  let c81_sxdd_dauKy = 0, c81_sxdd_cuoiKy = 0;
  let c82_thanhPham_dauKy = 0, c82_thanhPham_cuoiKy = 0;
  let c83_hangGui_dauKy = 0, c83_hangGui_cuoiKy = 0;

  const rC81 = findRow(/8\.1|sản\s*xuất\s*dở\s*dang/i);
  if (rC81 >= 0) {
    const nums = getRowNums(rC81);
    if (nums.length >= 2) { c81_sxdd_dauKy = nums[0]; c81_sxdd_cuoiKy = nums[1]; }
    else if (nums.length === 1) c81_sxdd_cuoiKy = nums[0];
  }
  if (c81_sxdd_cuoiKy === 0) {
    c81_sxdd_dauKy = parseNum(getCell(40, 15)) || parseNum(getCell(40, 16));
    c81_sxdd_cuoiKy = parseNum(getCell(40, 26)) || parseNum(getCell(40, 27));
  }

  const rC82 = findRow(/8\.2|thành\s*phẩm\s*tồn\s*kho/i);
  if (rC82 >= 0) {
    const nums = getRowNums(rC82);
    if (nums.length >= 2) { c82_thanhPham_dauKy = nums[0]; c82_thanhPham_cuoiKy = nums[1]; }
    else if (nums.length === 1) c82_thanhPham_cuoiKy = nums[0];
  }
  if (c82_thanhPham_cuoiKy === 0) {
    c82_thanhPham_dauKy = parseNum(getCell(41, 15)) || parseNum(getCell(41, 16));
    c82_thanhPham_cuoiKy = parseNum(getCell(41, 26)) || parseNum(getCell(41, 27));
  }

  const rC83 = findRow(/8\.3|hàng\s*gửi\s*đi\s*bán/i);
  if (rC83 >= 0) {
    const nums = getRowNums(rC83);
    if (nums.length >= 2) { c83_hangGui_dauKy = nums[0]; c83_hangGui_cuoiKy = nums[1]; }
    else if (nums.length === 1) c83_hangGui_cuoiKy = nums[0];
  }
  if (c83_hangGui_cuoiKy === 0) {
    c83_hangGui_dauKy = parseNum(getCell(42, 15)) || parseNum(getCell(42, 16));
    c83_hangGui_cuoiKy = parseNum(getCell(42, 26)) || parseNum(getCell(42, 27));
  }

  // CÂU 9: TỒN KHO THEO SẢN PHẨM
  const cau9_tonKhoSanPham: Array<{
    stt: number;
    tenNganh: string;
    maIO: string;
    c1_sxdd_dau: number;
    c2_sxdd_cuoi: number;
    c3_tp_dau: number;
    c4_tp_cuoi: number;
    c5_gui_dau: number;
    c6_gui_cuoi: number;
  }> = [
    {
      stt: 1,
      tenNganh,
      maIO,
      c1_sxdd_dau: c81_sxdd_dauKy,
      c2_sxdd_cuoi: c81_sxdd_cuoiKy,
      c3_tp_dau: c82_thanhPham_dauKy,
      c4_tp_cuoi: c82_thanhPham_cuoiKy,
      c5_gui_dau: c83_hangGui_dauKy,
      c6_gui_cuoi: c83_hangGui_cuoiKy
    }
  ];

  // PHẦN II: CÂU 10 (SẢN PHẨM CHÍNH)
  const extractRowVal = (regex: RegExp, fallbackRow: number, fallbackCol = 20): number => {
    const r = findRow(regex);
    if (r >= 0) {
      const nums = getRowNums(r);
      if (nums.length > 0) return nums[0];
    }
    return parseNum(getCell(fallbackRow, fallbackCol)) || parseNum(getCell(fallbackRow, fallbackCol + 1));
  };

  const mainDtt = extractRowVal(/Doanh\s*thu\s*thuần\s*ngành\s*sản\s*phẩm\s*chính|DTT/i, 60, 20) || cau5_tongDoanhThuThuan;
  const mainGvhb = extractRowVal(/Giá\s*vốn\s*hàng\s*bán\s*ngành\s*sản\s*phẩm\s*chính|GVHB/i, 61, 20) || Math.round(mainDtt * 0.8);
  const mainCfbh = extractRowVal(/Chi\s*phí\s*bán\s*hàng|CFBH/i, 62, 20);
  const mainCfql = extractRowVal(/Chi\s*phí\s*quản\s*lý|CFQL/i, 63, 20);
  const mainProfit = extractRowVal(/223|Lợi\s*nhuận\s*từ\s*hoạt\s*động\s*SXKD/i, 64, 20);
  const mainInterest = extractRowVal(/224|Trả\s*lãi\s*tiền\s*vay/i, 65, 20);
  const mainNvl = extractRowVal(/nguyên\s*liệu.*vật\s*liệu\s*trực\s*tiếp|CFNVL/i, 66, 20);
  const mainLabor = extractRowVal(/nhân\s*công\s*trực\s*tiếp|CFNC/i, 67, 20);
  const mainOverhead = extractRowVal(/sản\s*xuất\s*chung|CFSXC/i, 68, 20);
  const mainMachine = extractRowVal(/máy\s*thi\s*công|CFMTC/i, 69, 20);

  // CÂU 16: NHÂN CÔNG
  const tienLuong = extractRowVal(/182|Tiền\s*lương/i, 96, 20) || mainLabor || 0;
  const bhxh = extractRowVal(/183|Bảo\s*hiểm\s*xã\s*hội/i, 97, 20);
  const bhyt = extractRowVal(/184|Bảo\s*hiểm\s*y\s*tế/i, 98, 20);
  const bhtn = extractRowVal(/185|Bảo\s*hiểm\s*thất\s*nghiệp/i, 99, 20);
  const bhConNguoi = extractRowVal(/186/i, 100, 20);
  const kpcd = extractRowVal(/187|Kinh\s*phí\s*công\s*đoàn/i, 101, 20);
  const chiTrucTiep = extractRowVal(/195/i, 102, 20);
  const chiTraKhac = extractRowVal(/196/i, 103, 20);

  // CÂU 17: KHẤU HAO (MÃ 225)
  const khauHao = extractRowVal(/225|Khấu\s*hao/i, 109, 20);

  // CÂU 15: BÓC TÁCH NVL THEO MÃ IO
  const cau15_nvlChiTiet: Array<{ maIO: string; moTa: string; c1_giaTri: number; c2_tyLeNK: number; c3_giaCong: number }> = [];
  const rC15 = findRow(/15\.\s*Chi\s*phí\s*nguyên\s*liệu/i);
  if (rC15 >= 0) {
    for (let r = rC15 + 2; r < Math.min(rows.length, rC15 + 25); r++) {
      const row = rows[r];
      if (!row) continue;
      const rowStr = row.map(String).join(" ");
      if (/16\.\s*Chi\s*phí\s*nhân\s*công/i.test(rowStr) || /TỔNG\s*CỘNG/i.test(rowStr)) break;

      let rowIo = "";
      let moTa = "";
      for (let c = 0; c < row.length; c++) {
        const s = String(row[c]).trim();
        if (/^\d{3}$/.test(s) && !rowIo) rowIo = s;
        else if (s.length > 3 && !moTa && !/^\d+$/.test(s)) moTa = s;
      }
      const nums = getRowNums(r);
      if (nums.length >= 1 && rowIo) {
        cau15_nvlChiTiet.push({
          maIO: rowIo,
          moTa: moTa || `Nguyên vật liệu mã ${rowIo}`,
          c1_giaTri: nums[0],
          c2_tyLeNK: nums.length > 1 ? nums[1] : 0,
          c3_giaCong: nums.length > 2 ? nums[2] : 0
        });
      }
    }
  }

  // Nếu file ĐTV chưa bóc tách Câu 15, tự động phân bổ theo định mức ngành
  if (cau15_nvlChiTiet.length === 0) {
    const baseNvl = mainNvl || Math.round(mainGvhb * 0.65);
    detected.materialRatios.forEach(mat => {
      cau15_nvlChiTiet.push({
        maIO: mat.maIO,
        moTa: mat.moTa,
        c1_giaTri: Math.round(baseNvl * mat.ratio),
        c2_tyLeNK: mat.tyLeNK || 0,
        c3_giaCong: 0
      });
    });
  }

  // CÂU 18: DỊCH VỤ MUA NGOÀI THEO MÃ IO
  const cau18_muaNgoaiChiPhiKhac: Array<{ maIO: string; moTa: string; c1_giaTri: number; c2_tyLeNK: number }> = [];
  const rC18 = findRow(/18\.\s*Chi\s*phí\s*dịch\s*vụ\s*mua\s*ngoài/i);
  if (rC18 >= 0) {
    for (let r = rC18 + 2; r < Math.min(rows.length, rC18 + 25); r++) {
      const row = rows[r];
      if (!row) continue;
      const rowStr = row.map(String).join(" ");
      if (/19\./i.test(rowStr) || /TỔNG/i.test(rowStr)) break;

      let rowIo = "";
      let moTa = "";
      for (let c = 0; c < row.length; c++) {
        const s = String(row[c]).trim();
        if (/^\d{3}$/.test(s) && !rowIo) rowIo = s;
        else if (s.length > 3 && !moTa && !/^\d+$/.test(s)) moTa = s;
      }
      const nums = getRowNums(r);
      if (nums.length >= 1 && rowIo) {
        cau18_muaNgoaiChiPhiKhac.push({
          maIO: rowIo,
          moTa: moTa || `Dịch vụ mua ngoài mã ${rowIo}`,
          c1_giaTri: nums[0],
          c2_tyLeNK: nums.length > 1 ? nums[1] : 0
        });
      }
    }
  }

  if (cau18_muaNgoaiChiPhiKhac.length === 0) {
    const baseService = (mainCfbh + mainCfql + mainOverhead) || Math.round(mainGvhb * 0.25);
    detected.serviceRatios.forEach(srv => {
      cau18_muaNgoaiChiPhiKhac.push({
        maIO: srv.maIO,
        moTa: srv.moTa,
        c1_giaTri: Math.round(baseService * srv.ratio),
        c2_tyLeNK: srv.tyLeNK || 0
      });
    });
  }

  // TÍNH TOÁN GO VÀ IC CHÍNH XÁC
  const deltaTonKho = (c81_sxdd_cuoiKy + c82_thanhPham_cuoiKy + c83_hangGui_cuoiKy) -
                      (c81_sxdd_dauKy + c82_thanhPham_dauKy + c83_hangGui_dauKy);
  const go = mainDtt + deltaTonKho;

  const tongC15 = cau15_nvlChiTiet.reduce((s, r) => s + r.c1_giaTri, 0);
  const tongC18 = cau18_muaNgoaiChiPhiKhac.reduce((s, r) => s + r.c1_giaTri, 0);
  const ic = (tongC15 > 0 ? tongC15 : mainNvl) + (tongC18 > 0 ? tongC18 : (mainCfbh + mainCfql + mainOverhead));

  const tyLeIC_GO = go > 0 ? ic / go : 0;
  let danhGiaIC: "bình thường" | "quá cao (>0.75 hoặc >0.85)" | "quá thấp (<0.3)" = "bình thường";
  if (tyLeIC_GO > 0.85) danhGiaIC = "quá cao (>0.75 hoặc >0.85)";
  else if (tyLeIC_GO < 0.3) danhGiaIC = "quá thấp (<0.3)";

  const form: Phieu01IODNForm = {
    thongTinDinhDanh: {
      tenDoanhNghiep: tenDN,
      maSoThue: mst,
      diaChi,
      tinhTP,
      soDienThoai,
      loaiHinhKinhTe,
      maNganhC5: maC5,
      maIO,
      tenNganhChinh: tenNganh,
      sectorId: detected.id,
      sectorExplanation: `Nạp trực tiếp từ file Excel của Điều tra viên [${fileName}]. Ngành nghề nhận diện: ${detected.name}.`
    },
    phan1: {
      cau5_tongDoanhThuThuan,
      cau6_doanhThuSanPham,
      cau7_thongTinKhac,
      cau8_hangTonKhoDN: {
        c81_sxdd_dauKy,
        c81_sxdd_cuoiKy,
        c82_thanhPham_dauKy,
        c82_thanhPham_cuoiKy,
        c83_hangGui_dauKy,
        c83_hangGui_cuoiKy,
        tong_dauKy: c81_sxdd_dauKy + c82_thanhPham_dauKy + c83_hangGui_dauKy,
        tong_cuoiKy: c81_sxdd_cuoiKy + c82_thanhPham_cuoiKy + c83_hangGui_cuoiKy
      },
      cau9_tonKhoSanPham
    },
    phan2: {
      cau10_sxkdChinh: {
        dtt: mainDtt,
        gvhb: mainGvhb,
        cfbh: mainCfbh,
        cfql: mainCfql,
        ma223_loiNhuan: mainProfit,
        ma224_laiVay: mainInterest,
        cfnvl: mainNvl,
        cfnc: mainLabor,
        cfsxc: mainOverhead,
        cfmtc: mainMachine,
        c2_dtt: 0,
        c2_gvhb: 0,
        c2_cfbh: 0,
        c2_cfql: 0,
        c2_loiNhuan: 0,
        c2_laiVay: 0,
        c2_cfnvl: 0,
        c2_cfnc: 0,
        c2_cfsxc: 0,
        c2_cfmtc: 0
      },
      cau11_troCap: [
        { maIO: "202", ten: "Trợ giá sản phẩm", soTien: 0 },
        { maIO: "203", ten: "Trợ cấp vận chuyển", soTien: 0 },
        { maIO: "204", ten: "Trợ cấp xuất nhập khẩu", soTien: 0 },
        { maIO: "205", ten: "Trợ cấp quỹ lương", soTien: 0 },
        { maIO: "206", ten: "Trợ cấp giảm ô nhiễm", soTien: 0 },
        { maIO: "207", ten: "Các loại trợ cấp khác", soTien: 0 }
      ],
      cau12_thueLePhi: [
        { maIO: "208", ten: "Thuế VAT phát sinh", soTien: Math.round(mainDtt * 0.03), duocHachToan: false },
        { maIO: "210", ten: "Thuế TTĐB nội địa", soTien: 0, duocHachToan: false },
        { maIO: "214", ten: "Thuế bảo vệ môi trường", soTien: 0, duocHachToan: false },
        { maIO: "216", ten: "Thuế sử dụng đất", soTien: 0, duocHachToan: true },
        { maIO: "217", ten: "Lệ phí sản phẩm", soTien: 0, duocHachToan: true },
        { maIO: "230", ten: "Thuế khác", soTien: 0, duocHachToan: false }
      ],
      cau13_tienThuong191: 0,
      cau14_thuongMai: {
        c1_dtt: 0,
        c2_muaVe: 0,
        c3_vanTai: 0,
        c4_tonKhoDau: 0,
        c5_tonKhoCuoi: 0,
        c6_phiTM: 0
      }
    },
    phan3: {
      cau15_nvlChiTiet,
      cau16_nhanCong: {
        ma182_tienLuong: tienLuong,
        ma183_bhxh: bhxh,
        ma184_bhyt: bhyt,
        ma185_bhtn: bhtn,
        ma186_bhConNguoi: bhConNguoi,
        ma187_kpcd: kpcd,
        ma195_chiTrucTiep: chiTrucTiep,
        ma196_chiTraKhac: chiTraKhac,
        tongCong: tienLuong + bhxh + bhyt + bhtn + kpcd + chiTrucTiep + chiTraKhac,
        gc_182: 0,
        gc_183: 0,
        gc_184: 0,
        gc_185: 0,
        gc_186: 0,
        gc_187: 0,
        gc_195: 0,
        gc_196: 0,
        gc_tong: 0
      },
      cau17_khauHao225: khauHao,
      cau18_muaNgoaiChiPhiKhac
    },
    phan4: {
      cau194_vayNganHan320: 0,
      cau194_vayDaiHan338: 0,
      cau191_coTroCap: false,
      cau192_coBaoLanh: false,
      cau193_coVayNN: false
    },
    tongHopIO: {
      ic,
      go,
      tyLeIC_GO: Number(tyLeIC_GO.toFixed(4)),
      danhGiaIC
    }
  };

  return { form };
}

// ==============================================================================
// 3c. HÀM TỰ ĐỘNG PHÁT HIỆN ĐỊNH DẠNG FILE EXCEL (BCTC HAY PHIẾU 01 HAY BATCH DN)
// ==============================================================================
export function detectExcelFileType(wb: XLSX.WorkBook): "phieu01" | "bctc" | "batch_dn" | "unknown" {
  const sheetNames = wb.SheetNames.map(s => s.toLowerCase());

  // 1. Kiểm tra Phiếu 01/IO-DN của ĐTV
  for (const name of sheetNames) {
    if (name.includes("phieu01") || name.includes("phieu_01") || name.includes("phiếu 01") || name.includes("01-io") || name.includes("01_io")) {
      return "phieu01";
    }
  }

  // 2. Kiểm tra các sheet đặc trưng của Báo cáo tài chính
  for (const name of sheetNames) {
    if (
      name.includes("b01") || name.includes("b02") || name.includes("b03") || name.includes("f01") ||
      name.includes("cdkt") || name.includes("kqkd") || name.includes("kqhđkd") || name.includes("cdtk") ||
      name.includes("lctt") || name.includes("cân đối") || name.includes("kinh doanh")
    ) {
      return "bctc";
    }
  }

  // 3. Quét nội dung văn bản trong sheet đầu tiên
  const firstSheet = wb.Sheets[wb.SheetNames[0]];
  if (firstSheet) {
    const textSample = (XLSX.utils.sheet_to_json(firstSheet, { header: 1 }) as any[][])
      .slice(0, 35)
      .flat()
      .map(v => String(v || "").toLowerCase())
      .join(" ");

    if (textSample.includes("phiếu 01") || textSample.includes("01/io-dn") || textSample.includes("01/dn-io") || textSample.includes("tổng doanh thu thuần ngành")) {
      return "phieu01";
    }
    if (
      textSample.includes("bảng cân đối kế toán") || textSample.includes("kết quả hoạt động kinh doanh") ||
      textSample.includes("cân đối tài khoản") || textSample.includes("doanh thu thuần") ||
      textSample.includes("giá vốn hàng bán") || textSample.includes("b01-dn") || textSample.includes("b02-dn")
    ) {
      return "bctc";
    }
    if (textSample.includes("mã số thuế") && textSample.includes("tên doanh nghiệp") && (textSample.includes("doanh thu") || textSample.includes("giá vốn"))) {
      return "batch_dn";
    }
  }

  return "bctc"; // Mặc định thử bóc tách BCTC nếu không chắc chắn
}

// ==============================================================================
// 3d. HÀM PHÂN TÍCH VÀ BÓC TÁCH FILE EXCEL BÁO CÁO TÀI CHÍNH (B01, B02, F01, CDKT, KQKD)
// ==============================================================================
export function parseBctcExcel(fileData: ArrayBuffer | Uint8Array, fileName: string): RawBctcData {
  const wb = XLSX.read(fileData, { type: "array", cellDates: true, cellFormula: false });

  // Bộ làm sạch số đa năng (hỗ trợ dấu phẩy, dấu chấm, ngoặc đơn âm, định dạng VN/US)
  const cleanNumber = (val: any): number => {
    if (typeof val === "number") return isNaN(val) ? 0 : val;
    if (!val) return 0;
    const s = String(val).replace(/[\r\n\t]/g, "").trim();
    if (s === "-" || s === "" || s === "X" || s === "x" || s === "null" || s === "None" || s === "—") return 0;
    const isNeg = s.startsWith("(") && s.endsWith(")");
    // Xử lý dấu phân cách
    let clean = s.replace(/[\(\)]/g, "").replace(/\s/g, "");
    if (clean.includes(".") && clean.includes(",")) {
      if (clean.indexOf(".") < clean.indexOf(",")) {
        // Định dạng EU/VN: 1.000.000,00
        clean = clean.replace(/\./g, "").replace(/,/g, ".");
      } else {
        // Định dạng US: 1,000,000.00
        clean = clean.replace(/,/g, "");
      }
    } else if (clean.includes(",")) {
      // Có thể là 1,000,000 (hàng nghìn US) hoặc 100,5 (thập phân VN)
      const parts = clean.split(",");
      if (parts.length > 2 || (parts.length === 2 && parts[1].length === 3)) {
        clean = clean.replace(/,/g, "");
      } else {
        clean = clean.replace(/,/g, ".");
      }
    } else if (clean.includes(".")) {
      // Có thể là 1.000.000 (hàng nghìn VN)
      const parts = clean.split(".");
      if (parts.length > 2 || (parts.length === 2 && parts[1].length === 3 && clean.length > 6)) {
        clean = clean.replace(/\./g, "");
      }
    }
    const n = parseFloat(clean);
    if (isNaN(n)) return 0;
    return isNeg ? -Math.abs(n) : n;
  };

  // Thu thập tất cả các sheet dưới dạng mảng 2 chiều
  const sheetsData: { [sheetName: string]: any[][] } = {};
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (ws) {
      sheetsData[name] = (XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any[][]) || [];
    }
  }

  // 1. TRÍCH XUẤT THÔNG TIN ĐỊNH DANH DOANH NGHIỆP (MST, Tên, Địa chỉ, Năm, Thông tư)
  let mst = "";
  let tenDoanhNghiep = "";
  let diaChi = "";
  let nam = 2025;
  let thongTu: "TT200" | "TT133" | "Khac" = "TT200";

  // Quét qua các sheet để tìm thông tin Header
  for (const sheetName of wb.SheetNames) {
    const rows = sheetsData[sheetName] || [];
    const scanLimit = Math.min(rows.length, 25);

    for (let r = 0; r < scanLimit; r++) {
      const row = rows[r];
      if (!row) continue;
      const rowText = row.map(c => String(c || "")).join(" ");
      const lowerRow = rowText.toLowerCase();

      // Kiểm tra Thông tư
      if (lowerRow.includes("133/2016") || lowerRow.includes("tt133") || lowerRow.includes("b01a - dnn") || lowerRow.includes("b01b - dnn")) {
        thongTu = "TT133";
      }

      // Kiểm tra Năm
      const namMatch = lowerRow.match(/năm\s*(202\d)/i) || lowerRow.match(/kỳ\s*kế\s*toán.*(202\d)/i) || lowerRow.match(/\b(202\d)\b/);
      if (namMatch && !namMatch[0].includes("200/") && !namMatch[0].includes("133/")) {
        const foundNam = parseInt(namMatch[1] || namMatch[0]);
        if (foundNam >= 2020 && foundNam <= 2030) {
          nam = foundNam;
        }
      }

      // Kiểm tra MST
      if (!mst && (lowerRow.includes("mã số thuế") || lowerRow.includes("mst") || lowerRow.includes("m.s.t"))) {
        for (const cell of row) {
          const val = String(cell || "").trim();
          const match = val.match(/\b\d{10}(-\d{3})?\b/);
          if (match) {
            mst = match[0];
            break;
          }
        }
      }

      // Kiểm tra Tên doanh nghiệp
      if (!tenDoanhNghiep) {
        if (lowerRow.includes("tên doanh nghiệp") || lowerRow.includes("tên đơn vị") || lowerRow.includes("đơn vị báo cáo") || lowerRow.includes("đơn vị:")) {
          for (const cell of row) {
            const val = String(cell || "").trim();
            if (val && !/tên\s*(doanh\s*nghiệp|đơn\s*vị)|đơn\s*vị/i.test(val) && val.length > 5) {
              tenDoanhNghiep = val;
              break;
            }
          }
        } else {
          // Hoặc dòng bắt đầu bằng CÔNG TY / DOANH NGHIỆP / HỢP TÁC XÃ
          for (const cell of row) {
            const val = String(cell || "").trim();
            if (/^(công\s*ty|doanh\s*nghiệp|cty|dntn|hợp\s*tác\s*xã)\b/i.test(val) && val.length > 8) {
              tenDoanhNghiep = val;
              break;
            }
          }
        }
      }

      // Kiểm tra Địa chỉ
      if (!diaChi && (lowerRow.includes("địa chỉ") || lowerRow.includes("trụ sở"))) {
        for (const cell of row) {
          const val = String(cell || "").trim();
          if (val && !/địa\s*chỉ|trụ\s*sở/i.test(val) && val.length > 5) {
            diaChi = val;
            break;
          }
        }
      }
    }
  }

  // Fallback nếu chưa tìm thấy
  if (!tenDoanhNghiep) {
    // Thử lấy từ tên file (ví dụ BCTC_Cty_HongNam.xlsx)
    const cleanFileName = fileName.replace(/\.xlsx?$|\.xls$/i, "").replace(/bctc_?|báo_cáo_tài_chính_?/i, "").replace(/_/g, " ").trim();
    tenDoanhNghiep = cleanFileName ? `CÔNG TY ${cleanFileName.toUpperCase()}` : "Doanh nghiệp khảo sát IO (Nạp từ Excel BCTC)";
  }
  if (!mst) mst = "0100000000";
  if (!diaChi) diaChi = "Khu công nghiệp";

  // 2. KHỞI TẠO CÁC CHỈ TIÊU KẾT QUẢ KINH DOANH & CÂN ĐỐI KẾ TOÁN
  let dtt_10 = 0;
  let gvhb_11 = 0;
  let cfbh_25 = 0;
  let cfql_26 = 0;
  let laiVay_23 = 0;
  let ln_thuan_30 = 0;
  let ln_truoc_thue_50 = 0;

  let tonKho_141 = 0;
  let cfsxdd_154_dauKy = 0;
  let cfsxdd_154_cuoiKy = 0;
  let thanhPham_155_dauKy = 0;
  let thanhPham_155_cuoiKy = 0;
  let hangGuiBan_157_dauKy = 0;
  let hangGuiBan_157_cuoiKy = 0;
  let hangHoa_156_dauKy = 0;
  let hangHoa_156_cuoiKy = 0;

  let cf_nvl_621 = 0;
  let cf_nc_622 = 0;
  let cf_sxc_627 = 0;
  let cf_mtc_623 = 0;

  let khauHao_214 = 0;
  let thue_333 = 0;
  let tienLuong_334 = 0;
  let bhxh_338 = 0;
  let vayNganHan_320 = 0;
  let vayDaiHan_338 = 0;

  // Helper trích xuất số liệu của một hàng theo mã số hoặc tên chỉ tiêu
  // BCTC VN chuẩn: [Tên chỉ tiêu, Mã số, Thuyết minh, Năm nay/Kỳ này, Năm trước/Kỳ trước]
  const extractRowNumbers = (row: any[]): { namNay: number; namTruoc: number; allNums: number[] } => {
    const nums: { val: number; colIdx: number }[] = [];
    row.forEach((cell, idx) => {
      const n = cleanNumber(cell);
      if (n !== 0) nums.push({ val: n, colIdx: idx });
    });

    if (nums.length === 0) return { namNay: 0, namTruoc: 0, allNums: [] };
    if (nums.length === 1) return { namNay: nums[0].val, namTruoc: 0, allNums: [nums[0].val] };

    // Thường cột Năm nay đứng trước cột Năm trước
    return {
      namNay: nums[0].val,
      namTruoc: nums[1].val,
      allNums: nums.map(n => n.val)
    };
  };

  // Helper tìm hàng có chứa Mã số cụ thể (cột Mã số thường là 01, 10, 11, 25, 26, 140...)
  const findRowByCode = (rows: any[][], code: string): any[] | null => {
    for (const row of rows) {
      if (!row) continue;
      for (let c = 0; c < Math.min(row.length, 5); c++) {
        const val = String(row[c] || "").trim();
        if (val === code || val === `MS ${code}` || val === `Mã ${code}`) {
          return row;
        }
      }
    }
    return null;
  };

  // Helper tìm hàng theo Biểu thức chính quy (Regex) trên Tên chỉ tiêu
  const findRowByPattern = (rows: any[][], pattern: RegExp): any[] | null => {
    for (const row of rows) {
      if (!row) continue;
      const rowStr = row.slice(0, 4).map(c => String(c || "")).join(" ");
      if (pattern.test(rowStr)) {
        return row;
      }
    }
    return null;
  };

  // 3. QUÉT TOÀN BỘ CÁC SHEET TÌM KẾT QUẢ KINH DOANH (KQKD / B02-DN)
  for (const sheetName of wb.SheetNames) {
    const rows = sheetsData[sheetName] || [];

    // Tìm Doanh thu thuần (Mã 10)
    if (dtt_10 === 0) {
      const r10 = findRowByCode(rows, "10") || findRowByPattern(rows, /doanh\s*thu\s*thuần/i);
      if (r10) {
        const { namNay } = extractRowNumbers(r10);
        if (namNay > 0) dtt_10 = namNay;
      }
    }

    // Giá vốn hàng bán (Mã 11)
    if (gvhb_11 === 0) {
      const r11 = findRowByCode(rows, "11") || findRowByPattern(rows, /giá\s*vốn\s*hàng\s*bán/i);
      if (r11) {
        const { namNay } = extractRowNumbers(r11);
        if (namNay > 0) gvhb_11 = namNay;
      }
    }

    // Chi phí bán hàng (Mã 25)
    if (cfbh_25 === 0) {
      const r25 = findRowByCode(rows, "25") || findRowByPattern(rows, /chi\s*phí\s*bán\s*hàng/i);
      if (r25) {
        const { namNay } = extractRowNumbers(r25);
        if (namNay > 0) cfbh_25 = namNay;
      }
    }

    // Chi phí quản lý doanh nghiệp (Mã 26)
    if (cfql_26 === 0) {
      const r26 = findRowByCode(rows, "26") || findRowByPattern(rows, /chi\s*phí\s*quản\s*lý\s*doanh\s*nghiệp|chi\s*phí\s*quản\s*lý/i);
      if (r26) {
        const { namNay } = extractRowNumbers(r26);
        if (namNay > 0) cfql_26 = namNay;
      }
    }

    // Chi phí lãi vay (Mã 23)
    if (laiVay_23 === 0) {
      const r23 = findRowByCode(rows, "23") || findRowByPattern(rows, /lãi\s*tiền\s*vay|chi\s*phí\s*lãi\s*vay/i);
      if (r23) {
        const { namNay } = extractRowNumbers(r23);
        if (namNay > 0) laiVay_23 = namNay;
      }
    }

    // Lợi nhuận thuần (Mã 30)
    if (ln_thuan_30 === 0) {
      const r30 = findRowByCode(rows, "30") || findRowByPattern(rows, /lợi\s*nhuận\s*thuần\s*từ\s*hoạt\s*động\s*kinh\s*doanh/i);
      if (r30) {
        const { namNay } = extractRowNumbers(r30);
        ln_thuan_30 = namNay;
      }
    }

    // Tổng lợi nhuận trước thuế (Mã 50)
    if (ln_truoc_thue_50 === 0) {
      const r50 = findRowByCode(rows, "50") || findRowByPattern(rows, /tổng\s*lợi\s*nhuận\s*kế\s*toán\s*trước\s*thuế|lợi\s*nhuận\s*trước\s*thuế/i);
      if (r50) {
        const { namNay } = extractRowNumbers(r50);
        ln_truoc_thue_50 = namNay;
      }
    }

    // 4. QUÉT CÂN ĐỐI KẾ TOÁN (CĐKT / B01-DN) & HÀNG TỒN KHO
    // Hàng tồn kho (Mã 140 hoặc 141)
    if (tonKho_141 === 0) {
      const rTon = findRowByCode(rows, "141") || findRowByCode(rows, "140") || findRowByPattern(rows, /hàng\s*tồn\s*kho/i);
      if (rTon) {
        const { namNay } = extractRowNumbers(rTon);
        if (namNay > 0) tonKho_141 = namNay;
      }
    }

    // Chi phí SXKD dở dang (Mã 143 hoặc 154)
    if (cfsxdd_154_cuoiKy === 0 && cfsxdd_154_dauKy === 0) {
      const r154 = findRowByCode(rows, "143") || findRowByPattern(rows, /sản\s*xuất\s*kinh\s*doanh\s*dở\s*dang|sxkd\s*dở\s*dang|chi\s*phí\s*sản\s*xuất.*dở\s*dang/i);
      if (r154) {
        const { namNay, namTruoc } = extractRowNumbers(r154);
        cfsxdd_154_cuoiKy = namNay;
        cfsxdd_154_dauKy = namTruoc;
      }
    }

    // Thành phẩm (Mã 144 hoặc 155)
    if (thanhPham_155_cuoiKy === 0 && thanhPham_155_dauKy === 0) {
      const r155 = findRowByCode(rows, "144") || findRowByPattern(rows, /thành\s*phẩm\b/i);
      if (r155) {
        const { namNay, namTruoc } = extractRowNumbers(r155);
        thanhPham_155_cuoiKy = namNay;
        thanhPham_155_dauKy = namTruoc;
      }
    }

    // Hàng gửi đi bán (Mã 145 hoặc 157)
    if (hangGuiBan_157_cuoiKy === 0 && hangGuiBan_157_dauKy === 0) {
      const r157 = findRowByCode(rows, "145") || findRowByPattern(rows, /hàng\s*gửi\s*(đi\s*)?bán/i);
      if (r157) {
        const { namNay, namTruoc } = extractRowNumbers(r157);
        hangGuiBan_157_cuoiKy = namNay;
        hangGuiBan_157_dauKy = namTruoc;
      }
    }

    // Hàng hóa (Mã 146 hoặc 156)
    if (hangHoa_156_cuoiKy === 0 && hangHoa_156_dauKy === 0) {
      const r156 = findRowByCode(rows, "146") || findRowByPattern(rows, /hàng\s*hóa\b/i);
      if (r156) {
        const { namNay, namTruoc } = extractRowNumbers(r156);
        hangHoa_156_cuoiKy = namNay;
        hangHoa_156_dauKy = namTruoc;
      }
    }

    // Hao mòn TSCĐ / Khấu hao lũy kế (Mã 214)
    if (khauHao_214 === 0) {
      const r214 = findRowByCode(rows, "214") || findRowByPattern(rows, /hao\s*mòn\s*lũy\s*kế|khấu\s*hao\s*lũy\s*kế/i);
      if (r214) {
        const { namNay } = extractRowNumbers(r214);
        if (namNay !== 0) khauHao_214 = Math.abs(namNay);
      }
    }

    // Thuế và các khoản phải nộp Nhà nước (Mã 313)
    if (thue_333 === 0) {
      const r313 = findRowByCode(rows, "313") || findRowByPattern(rows, /thuế\s*và\s*các\s*khoản\s*phải\s*nộp/i);
      if (r313) {
        const { namNay } = extractRowNumbers(r313);
        if (namNay > 0) thue_333 = namNay;
      }
    }

    // Phải trả người lao động (Mã 314)
    if (tienLuong_334 === 0) {
      const r314 = findRowByCode(rows, "314") || findRowByPattern(rows, /phải\s*trả\s*người\s*lao\s*động/i);
      if (r314) {
        const { namNay } = extractRowNumbers(r314);
        if (namNay > 0) tienLuong_334 = namNay;
      }
    }

    // Phải trả khác / BHXH (Mã 319)
    if (bhxh_338 === 0) {
      const r319 = findRowByCode(rows, "319") || findRowByPattern(rows, /phải\s*trả\s*ngắn\s*hạn\s*khác/i);
      if (r319) {
        const { namNay } = extractRowNumbers(r319);
        if (namNay > 0) bhxh_338 = namNay;
      }
    }

    // Vay ngắn hạn (Mã 320)
    if (vayNganHan_320 === 0) {
      const r320 = findRowByCode(rows, "320") || findRowByPattern(rows, /vay.*ngắn\s*hạn/i);
      if (r320) {
        const { namNay } = extractRowNumbers(r320);
        if (namNay > 0) vayNganHan_320 = namNay;
      }
    }

    // Vay dài hạn (Mã 338)
    if (vayDaiHan_338 === 0) {
      const r338 = findRowByCode(rows, "338") || findRowByPattern(rows, /vay.*dài\s*hạn/i);
      if (r338) {
        const { namNay } = extractRowNumbers(r338);
        if (namNay > 0) vayDaiHan_338 = namNay;
      }
    }
  }

  // 5. QUÉT BẢNG CÂN ĐỐI TÀI KHOẢN (F01-DN / CDTK) NẾU CÓ ĐỂ LẤY CHI PHÍ CHI TIẾT
  for (const sheetName of wb.SheetNames) {
    const rows = sheetsData[sheetName] || [];
    for (const row of rows) {
      if (!row || row.length < 3) continue;
      const tkCode = String(row[0] || row[1] || "").trim();

      // TK 621: Chi phí nguyên vật liệu trực tiếp
      if (tkCode === "621" || tkCode === "TK621" || tkCode === "TK 621") {
        const nums = row.map(cleanNumber).filter(n => n > 0);
        if (nums.length > 0 && cf_nvl_621 === 0) cf_nvl_621 = nums[0];
      }
      // TK 622: Chi phí nhân công trực tiếp
      if (tkCode === "622" || tkCode === "TK622" || tkCode === "TK 622") {
        const nums = row.map(cleanNumber).filter(n => n > 0);
        if (nums.length > 0 && cf_nc_622 === 0) cf_nc_622 = nums[0];
      }
      // TK 627: Chi phí sản xuất chung
      if (tkCode === "627" || tkCode === "TK627" || tkCode === "TK 627") {
        const nums = row.map(cleanNumber).filter(n => n > 0);
        if (nums.length > 0 && cf_sxc_627 === 0) cf_sxc_627 = nums[0];
      }
      // TK 623: Chi phí máy thi công
      if (tkCode === "623" || tkCode === "TK623" || tkCode === "TK 623") {
        const nums = row.map(cleanNumber).filter(n => n > 0);
        if (nums.length > 0 && cf_mtc_623 === 0) cf_mtc_623 = nums[0];
      }
      // TK 641 hoặc 6421: Chi phí bán hàng
      if ((tkCode === "641" || tkCode === "6421") && cfbh_25 === 0) {
        const nums = row.map(cleanNumber).filter(n => n > 0);
        if (nums.length > 0) cfbh_25 = nums[0];
      }
      // TK 642 hoặc 6422: Chi phí quản lý DN
      if ((tkCode === "642" || tkCode === "6422") && cfql_26 === 0) {
        const nums = row.map(cleanNumber).filter(n => n > 0);
        if (nums.length > 0) cfql_26 = nums[0];
      }
      // TK 511: Doanh thu bán hàng & CCDV
      if (tkCode === "511" && dtt_10 === 0) {
        const nums = row.map(cleanNumber).filter(n => n > 0);
        if (nums.length > 0) dtt_10 = nums[0];
      }
      // TK 632: Giá vốn hàng bán
      if (tkCode === "632" && gvhb_11 === 0) {
        const nums = row.map(cleanNumber).filter(n => n > 0);
        if (nums.length > 0) gvhb_11 = nums[0];
      }
    }
  }

  // Fallback các chỉ tiêu cơ bản nếu file Excel chỉ có 1 số dòng tóm tắt
  if (dtt_10 === 0 && gvhb_11 > 0) {
    dtt_10 = Math.round(gvhb_11 * 1.15); // Ước tính nếu chỉ có giá vốn
  }
  if (gvhb_11 === 0 && dtt_10 > 0) {
    gvhb_11 = Math.round(dtt_10 * 0.82); // Ước tính nếu chỉ có doanh thu
  }
  if (tienLuong_334 === 0 && cf_nc_622 > 0) {
    tienLuong_334 = cf_nc_622;
  }
  if (bhxh_338 === 0 && tienLuong_334 > 0) {
    bhxh_338 = Math.round(tienLuong_334 * 0.235);
  }

  const rawBctc: RawBctcData = {
    mst,
    tenDoanhNghiep,
    diaChi,
    nam,
    thongTu,
    dtt_10,
    gvhb_11,
    cfbh_25,
    cfql_26,
    laiVay_23,
    ln_thuan_30,
    ln_truoc_thue_50,
    tonKho_141,
    cfsxdd_154_dauKy,
    cfsxdd_154_cuoiKy,
    thanhPham_155_dauKy,
    thanhPham_155_cuoiKy,
    hangGuiBan_157_dauKy,
    hangGuiBan_157_cuoiKy,
    hangHoa_156_dauKy,
    hangHoa_156_cuoiKy,
    cf_nvl_621,
    cf_nc_622,
    cf_sxc_627,
    cf_mtc_623,
    khauHao_214,
    thue_333,
    tienLuong_334,
    bhxh_338,
    vayNganHan_320,
    vayDaiHan_338,
    sourceFileName: fileName,
    sourceType: "excel"
  };

  return rawBctc;
}

// 3e. HÀM TẠO DỮ LIỆU MẪU ĐIỂN HÌNH TỪ PHIẾU ĐTV (Phieu_01_1001275111.xlsx) CÓ SẴN CÁC LỖI THỰC TẾ
export function createSampleDtvPhieu01(): Phieu01IODNForm {
  const dtt = 35400000000;
  const gvhb = 36800000000; // Lỗi: GVHB > DTT
  const cfbh = 1200000000;
  const cfql = 900000000;
  const loiNhuan = -2500000000;
  const laiVay = 13500000000; // Lỗi: Lãi vay / DTT = 38.1% (> 30%)
  const cfnvl = 24500000000;
  const cfnc = 8200000000;
  const cfsxc = 4100000000;
  const cfmtc = 0;

  const sxdd_dau = 1400000000;
  const sxdd_cuoi = 1850000000;
  const tp_dau = 850000000;
  const tp_cuoi = 920000000;
  const gui_dau = 0;
  const gui_cuoi = 0;

  const go = dtt + ((sxdd_cuoi + tp_cuoi + gui_cuoi) - (sxdd_dau + tp_dau + gui_dau));
  const ic = cfnvl + cfsxc + cfbh + cfql;
  const tyLeIC_GO = ic / go;

  return {
    thongTinDinhDanh: {
      tenDoanhNghiep: "CÔNG TY TNHH SẢN XUẤT THƯƠNG MẠI DỊCH VỤ CƯỜNG THỊNH",
      maSoThue: "1001275111",
      diaChi: "Thôn Thọ Lão, Xã Quang Hưng, Huyện Phù Cừ, Tỉnh Hưng Yên",
      tinhTP: "Hưng Yên",
      soDienThoai: "02213854999",
      loaiHinhKinhTe: "2",
      maNganhC5: "14100",
      maIO: "045",
      tenNganhChinh: "May mặc, sản xuất trang phục",
      sectorId: "det_may",
      sectorExplanation: "Mẫu phiếu điều tra viên thực tế (Phieu_01_1001275111.xlsx) có chứa các lỗi điển hình của ĐTV để phần mềm kiểm tra."
    },
    phan1: {
      cau5_tongDoanhThuThuan: dtt,
      cau6_doanhThuSanPham: [
        {
          stt: 1,
          tenNganh: "May mặc, sản xuất trang phục",
          maIO: "045",
          c1_dtt: dtt,
          c2_coGiaCong: 1, // Lỗi: Chọn có gia công
          c3_phiGiaCong: 0 // Lỗi: Phí gia công = 0
        }
      ],
      cau7_thongTinKhac: [
        {
          stt: 1,
          tenNganh: "May mặc, sản xuất trang phục",
          maIO: "045",
          c1_gvhb: gvhb,
          c2_gvChuyenBan: 0,
          c3_chiHo: 0,
          c4_chiThuong: 0
        }
      ],
      cau8_hangTonKhoDN: {
        c81_sxdd_dauKy: sxdd_dau,
        c81_sxdd_cuoiKy: sxdd_cuoi,
        c82_thanhPham_dauKy: tp_dau,
        c82_thanhPham_cuoiKy: tp_cuoi,
        c83_hangGui_dauKy: gui_dau,
        c83_hangGui_cuoiKy: gui_cuoi,
        tong_dauKy: sxdd_dau + tp_dau + gui_dau,
        tong_cuoiKy: sxdd_cuoi + tp_cuoi + gui_cuoi
      },
      cau9_tonKhoSanPham: [
        {
          stt: 1,
          tenNganh: "May mặc, sản xuất trang phục",
          maIO: "045",
          c1_sxdd_dau: sxdd_dau,
          c2_sxdd_cuoi: sxdd_cuoi,
          c3_tp_dau: tp_dau,
          c4_tp_cuoi: tp_cuoi,
          c5_gui_dau: gui_dau,
          c6_gui_cuoi: gui_cuoi
        }
      ]
    },
    phan2: {
      cau10_sxkdChinh: {
        dtt,
        gvhb,
        cfbh,
        cfql,
        ma223_loiNhuan: loiNhuan,
        ma224_laiVay: laiVay,
        cfnvl,
        cfnc,
        cfsxc,
        cfmtc,
        c2_dtt: 0,
        c2_gvhb: 0,
        c2_cfbh: 0,
        c2_cfql: 0,
        c2_loiNhuan: 0,
        c2_laiVay: 0,
        c2_cfnvl: 0,
        c2_cfnc: 0,
        c2_cfsxc: 0,
        c2_cfmtc: 0
      },
      cau11_troCap: [
        { maIO: "202", ten: "Trợ giá sản phẩm", soTien: 0 },
        { maIO: "203", ten: "Trợ cấp vận chuyển", soTien: 0 },
        { maIO: "204", ten: "Trợ cấp xuất nhập khẩu", soTien: 0 },
        { maIO: "205", ten: "Trợ cấp quỹ lương", soTien: 0 },
        { maIO: "206", ten: "Trợ cấp giảm ô nhiễm", soTien: 0 },
        { maIO: "207", ten: "Các loại trợ cấp khác", soTien: 0 }
      ],
      cau12_thueLePhi: [
        { maIO: "208", ten: "Thuế VAT phát sinh", soTien: 0, duocHachToan: false }, // Lỗi: Thuế = 0
        { maIO: "210", ten: "Thuế TTĐB nội địa", soTien: 0, duocHachToan: false },
        { maIO: "214", ten: "Thuế bảo vệ môi trường", soTien: 0, duocHachToan: false },
        { maIO: "216", ten: "Thuế sử dụng đất", soTien: 0, duocHachToan: true },
        { maIO: "217", ten: "Lệ phí liên quan sản phẩm chính", soTien: 0, duocHachToan: true },
        { maIO: "230", ten: "Thuế khác", soTien: 0, duocHachToan: false }
      ],
      cau13_tienThuong191: 0,
      cau14_thuongMai: {
        c1_dtt: 0,
        c2_muaVe: 0,
        c3_vanTai: 0,
        c4_tonKhoDau: 0,
        c5_tonKhoCuoi: 0,
        c6_phiTM: 0
      }
    },
    phan3: {
      cau15_nvlChiTiet: [
        { maIO: "044", moTa: "Vải dệt thoi, vải sợi tự nhiên và tổng hợp", c1_giaTri: 15500000000, c2_tyLeNK: 60, c3_giaCong: 0 },
        { maIO: "043", moTa: "Sợi dệt, chỉ may các loại", c1_giaTri: 4800000000, c2_tyLeNK: 30, c3_giaCong: 0 },
        { maIO: "082", moTa: "Phụ liệu kim loại (khóa kéo, cúc bấm, móc kim loại)", c1_giaTri: 2200000000, c2_tyLeNK: 40, c3_giaCong: 0 },
        { maIO: "055", moTa: "Bao bì carton, tem mác may mặc", c1_giaTri: 2000000000, c2_tyLeNK: 0, c3_giaCong: 0 }
      ],
      cau16_nhanCong: {
        ma182_tienLuong: cfnc,
        ma183_bhxh: 0, // Lỗi: Có lương nhưng không có BHXH
        ma184_bhyt: 0, // Lỗi: Không có BHYT
        ma185_bhtn: 0,
        ma186_bhConNguoi: 0,
        ma187_kpcd: 0, // Lỗi: Không có KPCĐ
        ma195_chiTrucTiep: 0,
        ma196_chiTraKhac: 0,
        tongCong: cfnc,
        gc_182: 0,
        gc_183: 0,
        gc_184: 0,
        gc_185: 0,
        gc_186: 0,
        gc_187: 0,
        gc_195: 0,
        gc_196: 0,
        gc_tong: 0
      },
      cau17_khauHao225: 0, // Lỗi: DN may mặc quy mô lớn nhưng khấu hao = 0
      cau18_muaNgoaiChiPhiKhac: [
        { maIO: "115", moTa: "Điện năng phục vụ xưởng may", c1_giaTri: 0, c2_tyLeNK: 0 }, // Lỗi: Điện = 0
        { maIO: "126", moTa: "Cước vận tải đường bộ chở hàng may", c1_giaTri: 3200000000, c2_tyLeNK: 0 },
        { maIO: "154", moTa: "Thuê nhà xưởng, kho bãi may mặc", c1_giaTri: 1800000000, c2_tyLeNK: 0 },
        { maIO: "179", moTa: "Bảo dưỡng, sửa chữa máy may công nghiệp", c1_giaTri: 1200000000, c2_tyLeNK: 0 }
      ]
    },
    phan4: {
      cau194_vayNganHan320: 15000000000,
      cau194_vayDaiHan338: 5000000000,
      cau191_coTroCap: false,
      cau192_coBaoLanh: false,
      cau193_coVayNN: false
    },
    tongHopIO: {
      ic,
      go,
      tyLeIC_GO: Number(tyLeIC_GO.toFixed(4)),
      danhGiaIC: tyLeIC_GO > 0.85 ? "quá cao (>0.75 hoặc >0.85)" : "bình thường"
    }
  };
}

// 4. BỘ KIỂM TRA LỖI & CẢNH BÁO THEO ĐÚNG TÀI LIỆU HƯỚNG DẪN PHIẾU 01/IO-DN
export interface AuditViolation {
  id: string;
  cau: string;
  type: "error" | "warning";
  tenLoi: string;
  moTa: string;
  giaTriThucTe: string;
  huongDanXuLy: string;
}

export function auditPhieu01IODN(form: Phieu01IODNForm): AuditViolation[] {
  const violations: AuditViolation[] = [];

  // LỖI 1: Ô có giá trị < 0 (trừ mã IO 223 - LN ở câu 10)
  if (form.phan1.cau5_tongDoanhThuThuan < 0) {
    violations.push({
      id: "v_dtt_am",
      cau: "Câu 5",
      type: "error",
      tenLoi: "Doanh thu thuần nhỏ hơn 0",
      moTa: "Các ô thông tin có giá trị <0 (trừ mã IO 223, câu 10)",
      giaTriThucTe: form.phan1.cau5_tongDoanhThuThuan.toLocaleString("vi-VN"),
      huongDanXuLy: "Kiểm tra lại số liệu doanh thu thuần phát sinh trên TK 511."
    });
  }

  // LỖI 2: Giá trị >= 100 nghìn tỷ
  if (form.phan1.cau5_tongDoanhThuThuan >= 100000000000000) {
    violations.push({
      id: "v_dtt_qua_lon",
      cau: "Câu 5",
      type: "error",
      tenLoi: "Giá trị quá lớn (>= 100 nghìn tỷ)",
      moTa: "Các ô thông tin có giá trị >=100.000.000.000.000 đồng",
      giaTriThucTe: form.phan1.cau5_tongDoanhThuThuan.toLocaleString("vi-VN"),
      huongDanXuLy: "Kiểm tra xem đơn vị tính có bị nhầm từ đồng sang nghìn đồng hay không."
    });
  }

  // LỖI 3: Câu 5 vs Câu 6 (Tổng DTT sản phẩm khác tổng DTT doanh nghiệp)
  const tongDTT_C6 = form.phan1.cau6_doanhThuSanPham.reduce((sum, sp) => sum + sp.c1_dtt, 0);
  if (Math.abs(tongDTT_C6 - form.phan1.cau5_tongDoanhThuThuan) > 1000) {
    violations.push({
      id: "v_c5_c6_lech",
      cau: "Câu 5 vs 6",
      type: "error",
      tenLoi: "Tổng doanh thu sản phẩm khác tổng doanh thu DN",
      moTa: "Tổng cộng doanh thu thuần của các sản phẩm KHÁC tổng doanh thu thuần của DN câu 5",
      giaTriThucTe: `Câu 5: ${form.phan1.cau5_tongDoanhThuThuan.toLocaleString("vi-VN")} | Câu 6: ${tongDTT_C6.toLocaleString("vi-VN")}`,
      huongDanXuLy: "Đối chiếu lại số liệu chia theo từng ngành sản phẩm IO với tổng TK 511."
    });
  }

  // CẢNH BÁO 4: DTT ngành chọn mẫu < 100 triệu
  const spChinh = form.phan1.cau6_doanhThuSanPham[0];
  if (spChinh && spChinh.c1_dtt > 0 && spChinh.c1_dtt < 100000000) {
    violations.push({
      id: "v_dtt_duoi_100tr",
      cau: "Câu 6",
      type: "warning",
      tenLoi: "Doanh thu ngành IO chọn mẫu < 100 triệu",
      moTa: "Nếu doanh thu thuần ngành I/O chọn mẫu < 100 triệu -> kết thúc điều tra, hiện thông báo và cảm ơn doanh nghiệp",
      giaTriThucTe: `${spChinh.c1_dtt.toLocaleString("vi-VN")} đồng`,
      huongDanXuLy: "Đề nghị ĐTV kiểm tra xem doanh nghiệp có còn hoạt động hoặc thay đổi ngành nghề chính không."
    });
  }

  // LỖI 5 & 6: Hoạt động gia công
  form.phan1.cau6_doanhThuSanPham.forEach((sp, idx) => {
    if (sp.c3_phiGiaCong > sp.c1_dtt && sp.c1_dtt > 0) {
      violations.push({
        id: `v_giacong_lon_dtt_${idx}`,
        cau: "Câu 6 Cột 3",
        type: "error",
        tenLoi: `Phí gia công lớn hơn doanh thu thuần (${sp.tenNganh})`,
        moTa: "Phí gia công (cột 3) > Doanh thu thuần sản phẩm (cột 1)",
        giaTriThucTe: `Phí GC: ${sp.c3_phiGiaCong.toLocaleString("vi-VN")} > DTT: ${sp.c1_dtt.toLocaleString("vi-VN")}`,
        huongDanXuLy: "Sửa lại phí gia công, không được vượt quá tổng doanh thu thuần."
      });
    }

    if (sp.c2_coGiaCong === 1 && sp.c3_phiGiaCong <= 0) {
      violations.push({
        id: `v_giacong_bang_0_${idx}`,
        cau: "Câu 6",
        type: "error",
        tenLoi: `Phí gia công = 0 khi có hoạt động gia công (${sp.tenNganh})`,
        moTa: "Phí gia công = 0 khi chọn có hoạt động gia công",
        giaTriThucTe: `Có hoạt động GC (Cột 2 = Có) nhưng Phí GC (Cột 3) = 0`,
        huongDanXuLy: "Điền doanh thu thu từ phí gia công vào cột 3."
      });
    }
  });

  // LỖI 7: Câu 7 theo Mã IO đặc thù
  form.phan1.cau7_thongTinKhac.forEach((sp, idx) => {
    if (["115", "116", "129", "130", "153"].includes(sp.maIO)) {
      if (sp.c1_gvhb <= 0) {
        violations.push({
          id: `v_gvhb_am_${idx}`,
          cau: "Câu 7 Cột 1",
          type: "error",
          tenLoi: `Mã IO ${sp.maIO} có trị giá vốn hàng bán <= 0`,
          moTa: "Mã IO 115, 116, 129, 130, 153 có trị giá vốn hàng bán <=0",
          giaTriThucTe: sp.c1_gvhb.toLocaleString("vi-VN"),
          huongDanXuLy: "Kiểm tra lại giá vốn hàng bán tương ứng TK 632."
        });
      } else {
        const dttTuongUng = form.phan1.cau6_doanhThuSanPham.find(s => s.maIO === sp.maIO)?.c1_dtt || 0;
        if (dttTuongUng > 0) {
          const tyLe = (sp.c1_gvhb / dttTuongUng) * 100;
          if (tyLe > 90 || tyLe < 10) {
            violations.push({
              id: `v_tyle_gvhb_${idx}`,
              cau: "Câu 7 Cột 1",
              type: "warning",
              tenLoi: `Mã IO ${sp.maIO} có tỷ lệ Giá vốn/Doanh thu bất thường`,
              moTa: "Mã IO 115, 116, 129, 130, 153 có Tỷ lệ Trị giá vốn/doanh thu > 90% hoặc <10%",
              giaTriThucTe: `${tyLe.toFixed(1)}%`,
              huongDanXuLy: "Đề nghị ĐTV xác minh tỷ suất lợi nhuận gộp của doanh nghiệp."
            });
          }
        }
      }
    }
  });

  // LỖI 8: Câu 8 vs Câu 9 (Tồn kho DN vs Tồn kho theo ngành sản phẩm)
  const tong9_sxdd_dau = form.phan1.cau9_tonKhoSanPham.reduce((s, r) => s + r.c1_sxdd_dau, 0);
  const tong9_sxdd_cuoi = form.phan1.cau9_tonKhoSanPham.reduce((s, r) => s + r.c2_sxdd_cuoi, 0);
  if (
    Math.abs(tong9_sxdd_dau - form.phan1.cau8_hangTonKhoDN.c81_sxdd_dauKy) > 1000 ||
    Math.abs(tong9_sxdd_cuoi - form.phan1.cau8_hangTonKhoDN.c81_sxdd_cuoiKy) > 1000
  ) {
    violations.push({
      id: "v_c8_c9_tonkho",
      cau: "Câu 8 vs 9",
      type: "error",
      tenLoi: "Tổng tồn kho chi tiết (Câu 9) lệch với Tổng tồn kho DN (Câu 8)",
      moTa: "Tổng cộng chi phí SX dở dang; giá trị thành phẩm tồn kho; giá trị hàng gửi đi bán của các sản phẩm KHÁC của DN (câu 8)",
      giaTriThucTe: `Câu 8: ${form.phan1.cau8_hangTonKhoDN.tong_cuoiKy.toLocaleString("vi-VN")} | Câu 9: ${(
        tong9_sxdd_cuoi
      ).toLocaleString("vi-VN")}`,
      huongDanXuLy: "Cân đối lại tổng các dòng câu 9 bằng đúng các chỉ tiêu ở câu 8."
    });
  }

  // LỖI 9: Câu 10 (CFBH + CFQL <= 0)
  const cfbh = form.phan2.cau10_sxkdChinh.cfbh;
  const cfql = form.phan2.cau10_sxkdChinh.cfql;
  if (cfbh + cfql <= 0 && form.phan2.cau10_sxkdChinh.dtt > 0) {
    violations.push({
      id: "v_cfbh_cfql_0",
      cau: "Câu 10",
      type: "error",
      tenLoi: "Tổng Chi phí bán hàng + Chi phí quản lý <= 0",
      moTa: "CFBH + CFQL <= 0",
      giaTriThucTe: `CFBH: ${cfbh.toLocaleString("vi-VN")} | CFQL: ${cfql.toLocaleString("vi-VN")}`,
      huongDanXuLy: "Doanh nghiệp hoạt động bắt buộc phát sinh chi phí bán hàng hoặc quản lý doanh nghiệp (TK 641, 642)."
    });
  }

  // CẢNH BÁO 10: Lãi vay (Mã 224) / DTT > 30%
  const laiVay = form.phan2.cau10_sxkdChinh.ma224_laiVay;
  const dttChinh = form.phan2.cau10_sxkdChinh.dtt;
  if (dttChinh > 0 && (laiVay / dttChinh) > 0.3) {
    violations.push({
      id: "v_laivay_30",
      cau: "Câu 10 Mã 224",
      type: "warning",
      tenLoi: "Tỷ lệ trả lãi tiền vay / Doanh thu thuần quá lớn (>30%)",
      moTa: "Mã 224/doanh thu thuần sp chính (câu 6) >30%",
      giaTriThucTe: `${((laiVay / dttChinh) * 100).toFixed(1)}%`,
      huongDanXuLy: "Đề nghị ĐTV kiểm tra xem doanh nghiệp có đòn bẩy tài chính cao hay kê khai nhầm chi phí tài chính khác."
    });
  }

  // LỖI 11: Tiền lương (Mã 182) = 0
  const tienLuong = form.phan3.cau16_nhanCong.ma182_tienLuong;
  if (tienLuong <= 0 && dttChinh > 0) {
    violations.push({
      id: "v_tienluong_0",
      cau: "Câu 16 Mã 182",
      type: "error",
      tenLoi: "Doanh nghiệp không phát sinh chi phí tiền lương (Mã 182 = 0)",
      moTa: 'Mã 182 = 0 --> "DN không phát sinh chi phí tiền lương"',
      giaTriThucTe: "0 đồng",
      huongDanXuLy: "Doanh nghiệp có doanh thu phải có lao động và chi phí tiền lương, phụ cấp (TK 334)."
    });
  }

  // CẢNH BÁO 12: BHXH / BHYT / KPCĐ = 0
  const bhxh = form.phan3.cau16_nhanCong.ma183_bhxh;
  if (tienLuong > 0 && bhxh <= 0) {
    violations.push({
      id: "v_bhxh_0",
      cau: "Câu 16 Mã 183",
      type: "warning",
      tenLoi: "Doanh nghiệp không phát sinh chi phí BHXH (Mã 183 = 0)",
      moTa: 'Mã 183 = 0 --> "Kiểm tra lại DN không phát sinh chi phí BHXH"',
      giaTriThucTe: "0 đồng",
      huongDanXuLy: "Đề nghị ĐTV kiểm tra việc đóng BHXH bắt buộc cho người lao động."
    });
  }

  // LỖI 13: Khấu hao TSCĐ (Mã 225) = 0
  const khauHao = form.phan3.cau17_khauHao225;
  if (khauHao <= 0 && dttChinh > 0) {
    violations.push({
      id: "v_khauhao_0",
      cau: "Câu 17 Mã 225",
      type: "error",
      tenLoi: "Doanh nghiệp không phát sinh chi phí khấu hao (Mã 225 = 0)",
      moTa: 'DN ngành CN, XD có mã 225 = 0 --> "DN không phát sinh chi phí khấu hao"',
      giaTriThucTe: "0 đồng",
      huongDanXuLy: "Kiểm tra số phát sinh Có TK 214 trong năm của doanh nghiệp."
    });
  }

  // CẢNH BÁO 14: Tỷ lệ IC/GO sơ bộ bất thường
  if (form.tongHopIO.danhGiaIC !== "bình thường") {
    violations.push({
      id: "v_ic_go_bat_thuong",
      cau: "Tổng hợp IC/GO",
      type: "warning",
      tenLoi: `Tỷ lệ Chi phí trung gian trên Giá trị sản xuất (IC/GO) ${form.tongHopIO.danhGiaIC}`,
      moTa: "Kiểm tra tỷ lệ IC/GO sơ bộ theo hướng dẫn trang 9 & 45",
      giaTriThucTe: `IC: ${form.tongHopIO.ic.toLocaleString("vi-VN")} | GO: ${form.tongHopIO.go.toLocaleString(
        "vi-VN"
      )} | Tỷ lệ: ${(form.tongHopIO.tyLeIC_GO * 100).toFixed(1)}%`,
      huongDanXuLy: "Đề nghị ĐTV đối chiếu lại giữa chi phí trung gian (IC) và giá trị sản xuất (GO)."
    });
  }

  // LỖI 15: Giá vốn hàng bán > Doanh thu thuần ở Câu 10 (Sản phẩm chính)
  const gvhbChinh = form.phan2.cau10_sxkdChinh.gvhb;
  if (dttChinh > 0 && gvhbChinh > dttChinh) {
    violations.push({
      id: "v_c10_gvhb_lon_dtt",
      cau: "Câu 10 Dòng 2",
      type: "warning",
      tenLoi: "Giá vốn hàng bán lớn hơn Doanh thu thuần sản phẩm chính",
      moTa: "Giá vốn hàng bán (TK 632) > Doanh thu thuần (TK 511) -> Doanh nghiệp kinh doanh dưới giá vốn (lỗ gộp)",
      giaTriThucTe: `DTT: ${dttChinh.toLocaleString("vi-VN")} | GVHB: ${gvhbChinh.toLocaleString("vi-VN")}`,
      huongDanXuLy: "ĐTV cần yêu cầu doanh nghiệp giải trình nguyên nhân bán dưới giá vốn, thanh lý tồn kho hoặc kiểm tra lại việc bóc tách số liệu."
    });
  }

  // LỖI 16: Chi phí trực tiếp đối với ngành Công nghiệp & Xây dựng
  const cfnvlChinh = form.phan2.cau10_sxkdChinh.cfnvl;
  const cfncChinh = form.phan2.cau10_sxkdChinh.cfnc;
  const cfsxcChinh = form.phan2.cau10_sxkdChinh.cfsxc;
  const cfmtcChinh = form.phan2.cau10_sxkdChinh.cfmtc;
  const tongCpTrucTiep = cfnvlChinh + cfncChinh + cfsxcChinh + cfmtcChinh;

  const isIndustrialOrConst = /công\s*nghiệp|chế\s*biến|chế\s*tạo|xây\s*dựng|may\s*mặc|cơ\s*khí|kim\s*loại/i.test(
    form.thongTinDinhDanh.tenNganhChinh || ""
  );
  if (isIndustrialOrConst && dttChinh > 0 && tongCpTrucTiep <= 0) {
    violations.push({
      id: "v_cptructiep_congnghiep",
      cau: "Câu 10 Chi phí trực tiếp",
      type: "error",
      tenLoi: "Doanh nghiệp ngành sản xuất / xây dựng nhưng không phát sinh chi phí trực tiếp",
      moTa: "Tổng chi phí NVL + Nhân công + Sản xuất chung <= 0",
      giaTriThucTe: `Tổng chi phí trực tiếp: ${tongCpTrucTiep.toLocaleString("vi-VN")} đồng`,
      huongDanXuLy: "Doanh nghiệp sản xuất hoặc xây dựng bắt buộc phải có chi phí NVL trực tiếp (TK 621), nhân công trực tiếp (TK 622) hoặc chi phí SX dở dang (TK 154)."
    });
  }

  // CẢNH BÁO 17: Lợi nhuận âm sâu (Mã 223 lỗ vượt 50% doanh thu)
  const loiNhuan223 = form.phan2.cau10_sxkdChinh.ma223_loiNhuan;
  if (dttChinh > 0 && loiNhuan223 < 0 && Math.abs(loiNhuan223) > dttChinh * 0.5) {
    violations.push({
      id: "v_loinhuan_am_sau",
      cau: "Câu 10 Mã 223",
      type: "warning",
      tenLoi: "Lợi nhuận SXKD âm vượt quá 50% doanh thu thuần",
      moTa: "Mã 223 < -50% Doanh thu thuần sản phẩm chính",
      giaTriThucTe: `${loiNhuan223.toLocaleString("vi-VN")} đồng (${((loiNhuan223 / dttChinh) * 100).toFixed(1)}% DTT)`,
      huongDanXuLy: "Đề nghị ĐTV kiểm tra xác minh nguyên nhân thua lỗ đột biến hoặc kiểm tra dấu âm/dương của chỉ tiêu."
    });
  }

  // CẢNH BÁO 18: Tỷ lệ chi phí bán hàng + QLDN quá lớn (> 30% DTT)
  if (dttChinh > 0 && ((cfbh + cfql) / dttChinh) > 0.3) {
    violations.push({
      id: "v_tyle_cp_dtt_30",
      cau: "Câu 10 CFBH + CFQL",
      type: "warning",
      tenLoi: "Tỷ lệ Chi phí bán hàng và QLDN trên Doanh thu thuần vượt quá 30%",
      moTa: "(CFBH + CFQL) / DTT > 30%",
      giaTriThucTe: `${(((cfbh + cfql) / dttChinh) * 100).toFixed(1)}% DTT`,
      huongDanXuLy: "Đề nghị ĐTV kiểm tra các khoản mục chi phí quản lý hoặc bán hàng bất thường trong kỳ."
    });
  }

  // CẢNH BÁO 19: Thuế phát sinh ở Câu 12 = 0 khi có doanh thu
  const tongThueC12 = form.phan2.cau12_thueLePhi.reduce((s, r) => s + r.soTien, 0);
  if (dttChinh > 0 && tongThueC12 <= 0) {
    violations.push({
      id: "v_thue_bang_0",
      cau: "Câu 12 Thuế & Phí",
      type: "warning",
      tenLoi: "Doanh nghiệp có doanh thu nhưng không kê khai các khoản thuế, phí phát sinh",
      moTa: "Tổng các khoản thuế, phí phát sinh ở câu 12 = 0",
      giaTriThucTe: "0 đồng",
      huongDanXuLy: "Kiểm tra thuế GTGT đầu ra phát sinh (TK 3331) hoặc các khoản thuế đất, lệ phí môn bài đã nộp."
    });
  }

  // CẢNH BÁO 20: Tiền lương phát sinh nhưng BHYT = 0 hoặc KPCĐ = 0
  const bhyt = form.phan3.cau16_nhanCong.ma184_bhyt;
  const kpcd = form.phan3.cau16_nhanCong.ma187_kpcd;
  if (tienLuong > 0 && bhyt <= 0) {
    violations.push({
      id: "v_bhyt_0",
      cau: "Câu 16 Mã 184",
      type: "warning",
      tenLoi: "Doanh nghiệp có tiền lương nhưng không kê khai BHYT (Mã 184 = 0)",
      moTa: "Mã 184 = 0 khi có tiền lương phát sinh",
      giaTriThucTe: "0 đồng",
      huongDanXuLy: "ĐTV kiểm tra các khoản trích theo lương của doanh nghiệp (TK 3383, 3384)."
    });
  }
  if (tienLuong > 0 && kpcd <= 0) {
    violations.push({
      id: "v_kpcd_0",
      cau: "Câu 16 Mã 187",
      type: "warning",
      tenLoi: "Doanh nghiệp có tiền lương nhưng không kê khai Kinh phí công đoàn (Mã 187 = 0)",
      moTa: "Mã 187 = 0 khi có tiền lương phát sinh",
      giaTriThucTe: "0 đồng",
      huongDanXuLy: "ĐTV kiểm tra trích nộp kinh phí công đoàn (TK 3382, 2% quỹ lương)."
    });
  }

  // CẢNH BÁO 21: Câu 15 - Chi phí tiền điện (mã IO 115) = 0 đối với ngành công nghiệp / may mặc / cơ khí
  if (isIndustrialOrConst) {
    const chiPhiDien = form.phan3.cau18_muaNgoaiChiPhiKhac.find(s => s.maIO === "115")?.c1_giaTri || 0;
    if (chiPhiDien <= 0 && dttChinh > 0) {
      violations.push({
        id: "v_cau18_dien_0",
        cau: "Câu 18 Mã IO 115",
        type: "warning",
        tenLoi: "Doanh nghiệp công nghiệp/sản xuất nhưng không kê khai chi phí điện năng (Mã IO 115)",
        moTa: "Mã IO 115 (Điện năng) = 0 đối với cơ sở sản xuất, nhà máy, xưởng may",
        giaTriThucTe: "0 đồng",
        huongDanXuLy: "Cơ sở sản xuất máy móc, may mặc bắt buộc phải tiêu thụ điện. ĐTV cần kiểm tra hóa đơn tiền điện hoặc tiểu khoản TK 627/642."
      });
    }
  }

  // CẢNH BÁO 22: Chi phí NVL trực tiếp vượt quá Doanh thu thuần
  if (dttChinh > 0 && cfnvlChinh > dttChinh) {
    violations.push({
      id: "v_cau15_nvl_lon_dtt",
      cau: "Câu 15 vs Câu 6",
      type: "warning",
      tenLoi: "Chi phí nguyên vật liệu trực tiếp lớn hơn Doanh thu thuần sản phẩm chính",
      moTa: "CFNVL trực tiếp > DTT",
      giaTriThucTe: `CFNVL: ${cfnvlChinh.toLocaleString("vi-VN")} | DTT: ${dttChinh.toLocaleString("vi-VN")}`,
      huongDanXuLy: "Kiểm tra xem nguyên vật liệu mua vào có kết chuyển vào sản phẩm dở dang hay doanh nghiệp bị hao hụt nguyên vật liệu đột biến."
    });
  }

  // CẢNH BÁO 23: Sản phẩm chính thuộc nhóm thuế Tiêu thụ đặc biệt (Mã IO 055, 056, 058, 098, 101)
  const maIOChinh = form.thongTinDinhDanh.maIO;
  const isSpecialConsumptionTaxIo = ["055", "056", "058", "098", "101", "55", "56", "58", "98"].includes(maIOChinh);
  if (isSpecialConsumptionTaxIo && dttChinh > 0) {
    const thueTTDB = form.phan2.cau12_thueLePhi.find(t => /tiêu\s*thụ\s*đặc\s*biệt|ttđb/i.test(t.ten))?.soTien || 0;
    if (thueTTDB <= 0) {
      violations.push({
        id: "v_thue_ttdb_0",
        cau: "Câu 12 Thuế TTĐB",
        type: "warning",
        tenLoi: `Sản phẩm chính ngành IO ${maIOChinh} (Bia, Rượu, Thuốc lá, Ô tô, Xe máy) phải có thuế TTĐB`,
        moTa: "Sản phẩm chính mã IO 55, 56, 58, 98, 101 có thuế TTĐB nhưng chưa kê khai",
        giaTriThucTe: "0 đồng",
        huongDanXuLy: "Kiểm tra số phát sinh thuế Tiêu thụ đặc biệt (TK 3332) của doanh nghiệp trong năm."
      });
    }
  }

  // CẢNH BÁO 24: Thuế phát sinh / DTT > 30%
  if (dttChinh > 0 && (tongThueC12 / dttChinh) > 0.3) {
    violations.push({
      id: "v_thue_tren_30_dtt",
      cau: "Câu 12 Thuế & Phí",
      type: "warning",
      tenLoi: "Tỷ lệ Thuế, phí phát sinh trên Doanh thu thuần vượt quá 30%",
      moTa: "Tổng thuế phát sinh / Doanh thu thuần (Câu 6) > 30%",
      giaTriThucTe: `${((tongThueC12 / dttChinh) * 100).toFixed(1)}%`,
      huongDanXuLy: "Đề nghị ĐTV kiểm tra xem doanh nghiệp có kê khai gộp thuế TNDN hoãn lại hoặc thuế xuất nhập khẩu đột biến hay không."
    });
  }

  // CẢNH BÁO 25: Doanh nghiệp đơn ngành (chỉ có 1 sản phẩm chính ở câu 6)
  if (form.phan1.cau6_doanhThuSanPham.length === 1 && dttChinh > 0) {
    const dttTong = form.phan1.cau5_tongDoanhThuThuan;
    if (Math.abs(dttChinh - dttTong) > 1000) {
      violations.push({
        id: "v_don_nganh_dtt_lech",
        cau: "Câu 6 vs Câu 5",
        type: "error",
        tenLoi: "Doanh nghiệp đơn ngành nhưng Doanh thu sản phẩm chính khác Tổng DTT",
        moTa: "Nếu DN đơn ngành, DTT sản phẩm chính phải bằng đúng DTT của DN",
        giaTriThucTe: `DTT SP chính: ${dttChinh.toLocaleString("vi-VN")} | Tổng DTT: ${dttTong.toLocaleString("vi-VN")}`,
        huongDanXuLy: "Đối chiếu và sửa lại để Doanh thu sản phẩm chính bằng Tổng doanh thu thuần."
      });
    }
  }

  // CẢNH BÁO 26 (G34 & G36): Kiểm tra tính hợp lệ Cột 2 & Cột 3 ở Câu 15 (NVL)
  form.phan3.cau15_nvlChiTiet.forEach((item, idx) => {
    if (item.c2_tyLeNK > 100 || item.c2_tyLeNK < 0) {
      violations.push({
        id: `v_c15_tylenk_sai_${idx}`,
        cau: "Câu 15 Cột 2",
        type: "error",
        tenLoi: `Tỷ lệ nhập khẩu ở mã IO ${item.maIO} không hợp lệ (>100% hoặc <0%)`,
        moTa: "Cột 2 (Tỷ lệ nhập khẩu) > 100%",
        giaTriThucTe: `${item.c2_tyLeNK}%`,
        huongDanXuLy: "Điền tỷ lệ nhập khẩu từ 0% đến 100%."
      });
    }
    if (item.c3_giaCong > item.c1_giaTri && item.c1_giaTri > 0) {
      violations.push({
        id: `v_c15_giacong_lon_trigia_${idx}`,
        cau: "Câu 15 Cột 3",
        type: "error",
        tenLoi: `Chi phí nhận gia công lớn hơn trị giá chi phí ở mã IO ${item.maIO}`,
        moTa: "Cột 3 > Cột 1 ở câu 15",
        giaTriThucTe: `Gia công: ${item.c3_giaCong.toLocaleString("vi-VN")} > Trị giá: ${item.c1_giaTri.toLocaleString("vi-VN")}`,
        huongDanXuLy: "Trị giá nguyên vật liệu nhận gia công không được lớn hơn tổng giá trị."
      });
    }
  });

  // CẢNH BÁO 27 (G35): Có nhận gia công nhưng tổng Cột 3 ở Câu 15 = 0
  const hasGiaCongC6 = form.phan1.cau6_doanhThuSanPham.some(s => s.c2_coGiaCong === 1);
  const tongGiaCongC15 = form.phan3.cau15_nvlChiTiet.reduce((s, r) => s + r.c3_giaCong, 0);
  if (hasGiaCongC6 && tongGiaCongC15 <= 0 && dttChinh > 0) {
    violations.push({
      id: "v_c15_giacong_c3_0",
      cau: "Câu 15 Cột 3",
      type: "warning",
      tenLoi: "Doanh nghiệp có nhận gia công ở Câu 6 nhưng chi phí nhận gia công ở Câu 15 bằng 0",
      moTa: "Tổng cộng cột 3 = 0 khi sản phẩm chính ở câu 6 có hoạt động gia công",
      giaTriThucTe: "0 đồng",
      huongDanXuLy: "Kê khai chi phí nguyên vật liệu nhận gia công tương ứng vào Cột 3 của Câu 15."
    });
  }

  // CẢNH BÁO 28 (G38): Doanh nghiệp sản xuất/xây dựng nhưng chi phí NVL ở Câu 15 = 0
  const tongNvlC15 = form.phan3.cau15_nvlChiTiet.reduce((s, r) => s + r.c1_giaTri, 0);
  if (isIndustrialOrConst && tongNvlC15 <= 0 && dttChinh > 0) {
    violations.push({
      id: "v_c15_nvl_tong_0",
      cau: "Câu 15 Tổng cộng",
      type: "error",
      tenLoi: "Doanh nghiệp công nghiệp, xây dựng có tổng chi phí nguyên vật liệu bằng 0",
      moTa: "DN ngành CN, XD có NVL (dòng tổng cộng câu 15) = 0",
      giaTriThucTe: "0 đồng",
      huongDanXuLy: "Cơ sở sản xuất hoặc xây dựng bắt buộc phát sinh chi phí nguyên vật liệu (TK 621/154)."
    });
  }

  // CẢNH BÁO 29 (G50): Doanh nghiệp vận tải (Mã IO 131-138) nhưng xăng dầu (Mã IO 068) = 0
  const isTransport = ["131", "132", "133", "134", "135", "136", "137", "138"].includes(maIOChinh);
  if (isTransport && dttChinh > 0) {
    const cpXangDau = form.phan3.cau15_nvlChiTiet.find(r => r.maIO === "068" || r.maIO === "68")?.c1_giaTri || 0;
    if (cpXangDau <= 0) {
      violations.push({
        id: "v_vantai_xangdau_0",
        cau: "Câu 15 Mã IO 068",
        type: "error",
        tenLoi: "Doanh nghiệp vận tải nhưng không kê khai chi phí nhiên liệu, xăng dầu (Mã IO 068)",
        moTa: "DN vận tải (mã IO 131 đến 138) có mã 68 = 0",
        giaTriThucTe: "0 đồng",
        huongDanXuLy: "Doanh nghiệp vận tải đường bộ/thủy bắt buộc tiêu hao xăng dầu, mỡ nhờn. ĐTV cần rà soát lại."
      });
    }
  }

  // CẢNH BÁO 30 (G48): Chi phí nước sạch (Mã IO 118) = 0 đối với ngành công nghiệp
  if (isIndustrialOrConst && dttChinh > 500000000) {
    const cpNuoc = form.phan3.cau18_muaNgoaiChiPhiKhac.find(r => r.maIO === "118")?.c1_giaTri || 0;
    if (cpNuoc <= 0) {
      violations.push({
        id: "v_nuocsach_118_0",
        cau: "Câu 18 Mã IO 118",
        type: "warning",
        tenLoi: "Doanh nghiệp sản xuất nhưng không phát sinh chi phí nước sạch (Mã IO 118)",
        moTa: "Các DN sản xuất có mã 118 = 0",
        giaTriThucTe: "0 đồng",
        huongDanXuLy: "Kiểm tra hóa đơn nước sinh hoạt hoặc nước sản xuất công nghiệp."
      });
    }
  }

  // CẢNH BÁO 31 (G55 & G56): Dịch vụ viễn thông (146) hoặc tài chính ngân hàng (149) = 0 khi có doanh thu lớn
  if (dttChinh > 1000000000) {
    const cpVienThong = form.phan3.cau18_muaNgoaiChiPhiKhac.find(r => r.maIO === "146")?.c1_giaTri || 0;
    const cpNganHang = form.phan3.cau18_muaNgoaiChiPhiKhac.find(r => r.maIO === "149")?.c1_giaTri || 0;
    if (cpVienThong <= 0) {
      violations.push({
        id: "v_vienthong_146_0",
        cau: "Câu 18 Mã IO 146",
        type: "warning",
        tenLoi: "Doanh nghiệp có doanh thu lớn nhưng chi phí viễn thông, internet (Mã 146) = 0",
        moTa: "Mã 146 = 0 đối với DN có doanh thu > 1 tỷ đồng",
        giaTriThucTe: "0 đồng",
        huongDanXuLy: "Kiểm tra chi phí cước điện thoại, internet văn phòng."
      });
    }
    if (cpNganHang <= 0) {
      violations.push({
        id: "v_nganhang_149_0",
        cau: "Câu 18 Mã IO 149",
        type: "warning",
        tenLoi: "Doanh nghiệp có doanh thu lớn nhưng chi phí dịch vụ ngân hàng (Mã 149) = 0",
        moTa: "Mã 149 = 0 đối với DN có doanh thu > 1 tỷ đồng",
        giaTriThucTe: "0 đồng",
        huongDanXuLy: "Kiểm tra phí chuyển tiền, phí duy trì tài khoản ngân hàng (TK 6425/635)."
      });
    }
  }

  return violations;
}
