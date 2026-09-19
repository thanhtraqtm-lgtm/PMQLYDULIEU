import * as XLSX from "xlsx";
import { Phieu01IODNForm, AuditViolation } from "./bctcParser";

// Xuất file Mẫu Phiếu 01/IO-DN đầy đủ theo chuẩn Cục Thống Kê
export function exportPhieu01IODNToExcel(form: Phieu01IODNForm, violations: AuditViolation[]) {
  const wb = XLSX.utils.book_new();

  // === SHEET 1: PHIẾU 01/IO-DN ===
  const rows: any[][] = [];

  // Header Cục Thống Kê
  rows.push(["CỤC THỐNG KÊ", "", "", "PHIẾU SỐ: 01/IO-DN", "", ""]);
  rows.push(["ĐIỀU TRA LẬP BẢNG CÂN ĐỐI LIÊN NGÀNH VÀ TÍNH HỆ SỐ CHI PHÍ TRUNG GIAN (IO 2026)", "", "", "", "", ""]);
  rows.push(["PHIẾU THU THẬP THÔNG TIN VỀ DOANH NGHIỆP KHÔNG THUỘC LĨNH VỰC TÀI CHÍNH, NGÂN HÀNG, BẢO HIỂM", "", "", "", "", ""]);
  rows.push(["(Thực hiện theo Quyết định số 1124/QĐ-CTK ngày 11 tháng 9 năm 2025)", "", "", "", "", ""]);
  rows.push([]);

  // THÔNG TIN ĐỊNH DANH
  rows.push(["THÔNG TIN ĐỊNH DANH DOANH NGHIỆP", "", "", "", "", ""]);
  rows.push(["1. Tên doanh nghiệp:", form.thongTinDinhDanh.tenDoanhNghiep]);
  rows.push(["Mã số thuế:", form.thongTinDinhDanh.maSoThue]);
  rows.push(["2. Địa chỉ:", form.thongTinDinhDanh.diaChi]);
  rows.push(["Tỉnh/TP:", form.thongTinDinhDanh.tinhTP]);
  rows.push(["Số điện thoại:", form.thongTinDinhDanh.soDienThoai]);
  rows.push(["3. Loại hình kinh tế:", form.thongTinDinhDanh.loaiHinhKinhTe === "1" ? "1. Nhà nước" : form.thongTinDinhDanh.loaiHinhKinhTe === "3" ? "3. Có vốn đầu tư nước ngoài" : "2. Ngoài nhà nước"]);
  rows.push(["4. Ngành sản phẩm chính:", `${form.thongTinDinhDanh.tenNganhChinh} (Mã C5: ${form.thongTinDinhDanh.maNganhC5} - Mã IO: ${form.thongTinDinhDanh.maIO})`]);
  rows.push([]);

  // PHẦN I
  rows.push(["PHẦN I: KẾT QUẢ HOẠT ĐỘNG SẢN XUẤT KINH DOANH CỦA DOANH NGHIỆP NĂM 2025", "", "", "", "", ""]);
  rows.push(["5. Tổng doanh thu thuần hoạt động SXKD năm 2025 (đồng):", form.phan1.cau5_tongDoanhThuThuan]);
  rows.push([]);

  // Bảng Câu 6
  rows.push(["6. Doanh thu thuần chia theo sản phẩm SXKD của DN năm 2025", "", "", "", "", ""]);
  rows.push(["STT", "Tên ngành sản phẩm", "Mã IO", "Doanh thu thuần chia theo sản phẩm (1)", "DN có thực hiện gia công không? (2)", "Doanh thu từ thu phí gia công (3)"]);
  form.phan1.cau6_doanhThuSanPham.forEach(sp => {
    rows.push([
      sp.stt,
      sp.tenNganh,
      sp.maIO,
      sp.c1_dtt,
      sp.c2_coGiaCong === 1 ? "(1) Có" : "(2) Không",
      sp.c3_phiGiaCong
    ]);
  });
  const tongDttC6 = form.phan1.cau6_doanhThuSanPham.reduce((s, r) => s + r.c1_dtt, 0);
  const tongPhiGcC6 = form.phan1.cau6_doanhThuSanPham.reduce((s, r) => s + r.c3_phiGiaCong, 0);
  rows.push(["", "TỔNG CỘNG CÂU 6", "X", tongDttC6, "", tongPhiGcC6]);
  rows.push([]);

  // Bảng Câu 7
  rows.push(["7. Thông tin khác về sản phẩm SXKD của doanh nghiệp trong năm 2025", "", "", "", "", ""]);
  rows.push(["STT", "Tên ngành sản phẩm", "Mã IO", "Trị giá vốn hàng bán (1)", "Trị giá vốn hàng chuyển bán (2)", "Chi hộ khách hàng (3)", "Chi trả thưởng (4)"]);
  form.phan1.cau7_thongTinKhac.forEach(sp => {
    rows.push([sp.stt, sp.tenNganh, sp.maIO, sp.c1_gvhb, sp.c2_gvChuyenBan, sp.c3_chiHo, sp.c4_chiThuong]);
  });
  rows.push([]);

  // Bảng Câu 8
  rows.push(["8. Thông tin hàng tồn kho của DN năm 2025", "", "", "", "", ""]);
  rows.push(["Chỉ tiêu", "Thời điểm 01/01/2025 (Đầu kỳ)", "Thời điểm 31/12/2025 (Cuối kỳ)"]);
  rows.push(["8.1 Chi phí sản xuất dở dang (đồng)", form.phan1.cau8_hangTonKhoDN.c81_sxdd_dauKy, form.phan1.cau8_hangTonKhoDN.c81_sxdd_cuoiKy]);
  rows.push(["8.2 Giá trị thành phẩm tồn kho (đồng)", form.phan1.cau8_hangTonKhoDN.c82_thanhPham_dauKy, form.phan1.cau8_hangTonKhoDN.c82_thanhPham_cuoiKy]);
  rows.push(["8.3 Giá trị hàng gửi đi bán (đồng)", form.phan1.cau8_hangTonKhoDN.c83_hangGui_dauKy, form.phan1.cau8_hangTonKhoDN.c83_hangGui_cuoiKy]);
  rows.push(["TỔNG CỘNG CÂU 8", form.phan1.cau8_hangTonKhoDN.tong_dauKy, form.phan1.cau8_hangTonKhoDN.tong_cuoiKy]);
  rows.push([]);

  // PHẦN II
  rows.push(["PHẦN II: KẾT QUẢ HOẠT ĐỘNG SXKD NGÀNH SẢN PHẨM CHÍNH NĂM 2025", "", "", "", "", ""]);
  rows.push(["10. Thông tin về hoạt động SXKD SẢN PHẨM CHÍNH của DN năm 2025", "", "", "", "", ""]);
  rows.push(["Tên chỉ tiêu", "Mã I/O", "CỦA SẢN PHẨM CHÍNH (Cột 1)", "Trong đó: hoạt động gia công (Cột 2)"]);
  rows.push(["Doanh thu thuần ngành sản phẩm chính", "DTT", form.phan2.cau10_sxkdChinh.dtt, form.phan2.cau10_sxkdChinh.c2_dtt]);
  rows.push(["Giá vốn hàng bán ngành sản phẩm chính", "GVHB", form.phan2.cau10_sxkdChinh.gvhb, form.phan2.cau10_sxkdChinh.c2_gvhb]);
  rows.push(["Chi phí bán hàng phân bổ cho ngành SP chính", "CFBH", form.phan2.cau10_sxkdChinh.cfbh, form.phan2.cau10_sxkdChinh.c2_cfbh]);
  rows.push(["Chi phí quản lý phân bổ cho ngành SP chính", "CFQL", form.phan2.cau10_sxkdChinh.cfql, form.phan2.cau10_sxkdChinh.c2_cfql]);
  rows.push(["Lợi nhuận từ hoạt động SXKD ngành sản phẩm chính", "223", form.phan2.cau10_sxkdChinh.ma223_loiNhuan, form.phan2.cau10_sxkdChinh.c2_loiNhuan]);
  rows.push(["Trả lãi tiền vay ngân hàng cho hoạt động SXKD SP chính", "224", form.phan2.cau10_sxkdChinh.ma224_laiVay, form.phan2.cau10_sxkdChinh.c2_laiVay]);
  rows.push(["Chi phí nguyên liệu, vật liệu trực tiếp", "CFNVL", form.phan2.cau10_sxkdChinh.cfnvl, form.phan2.cau10_sxkdChinh.c2_cfnvl]);
  rows.push(["Chi phí nhân công trực tiếp", "CFNC", form.phan2.cau10_sxkdChinh.cfnc, form.phan2.cau10_sxkdChinh.c2_cfnc]);
  rows.push(["Chi phí sản xuất chung", "CFSXC", form.phan2.cau10_sxkdChinh.cfsxc, form.phan2.cau10_sxkdChinh.c2_cfsxc]);
  rows.push(["Chi phí sử dụng máy thi công (ngành XD)", "CFMTC", form.phan2.cau10_sxkdChinh.cfmtc, form.phan2.cau10_sxkdChinh.c2_cfmtc]);
  rows.push([]);

  // Bảng Câu 15 (Bóc tách Chi phí Nguyên vật liệu theo Mã IO)
  rows.push(["15. Chi phí nguyên liệu, vật liệu để SXKD SẢN PHẨM CHÍNH năm 2025 (theo Mã IO)", "", "", "", "", ""]);
  rows.push(["Mã I/O", "Tên nguyên liệu, vật liệu", "Trị giá (đồng) (1)", "Tỷ lệ nhập khẩu (%) (2)", "Trong đó: phục vụ gia công (3)"]);
  form.phan3.cau15_nvlChiTiet.forEach(item => {
    rows.push([item.maIO, item.moTa, item.c1_giaTri, `${item.c2_tyLeNK}%`, item.c3_giaCong]);
  });
  const tongC15 = form.phan3.cau15_nvlChiTiet.reduce((s, r) => s + r.c1_giaTri, 0);
  rows.push(["TỔNG", "TỔNG CỘNG CHI PHÍ NGUYÊN VẬT LIỆU (CÂU 15)", tongC15, "", 0]);
  rows.push([]);

  // Bảng Câu 16 (Nhân công)
  rows.push(["16. Chi phí nhân công để sản xuất kinh doanh SẢN PHẨM CHÍNH năm 2025", "", "", "", "", ""]);
  rows.push(["Chỉ tiêu", "Mã I/O", "Chi phí nhân công SP CHÍNH (Cột 1)", "Trong đó: hoạt động gia công (Cột 2)"]);
  rows.push(["- Tiền lương, tiền công và các khoản phụ cấp", "182", form.phan3.cau16_nhanCong.ma182_tienLuong, form.phan3.cau16_nhanCong.gc_182]);
  rows.push(["- Bảo hiểm xã hội", "183", form.phan3.cau16_nhanCong.ma183_bhxh, form.phan3.cau16_nhanCong.gc_183]);
  rows.push(["- Bảo hiểm y tế", "184", form.phan3.cau16_nhanCong.ma184_bhyt, form.phan3.cau16_nhanCong.gc_184]);
  rows.push(["- Bảo hiểm thất nghiệp", "185", form.phan3.cau16_nhanCong.ma185_bhtn, form.phan3.cau16_nhanCong.gc_185]);
  rows.push(["- Bảo hiểm con người", "186", form.phan3.cau16_nhanCong.ma186_bhConNguoi, form.phan3.cau16_nhanCong.gc_186]);
  rows.push(["- Kinh phí công đoàn", "187", form.phan3.cau16_nhanCong.ma187_kpcd, form.phan3.cau16_nhanCong.gc_187]);
  rows.push(["- Các khoản chi trực tiếp bằng tiền", "195", form.phan3.cau16_nhanCong.ma195_chiTrucTiep, form.phan3.cau16_nhanCong.gc_195]);
  rows.push(["- Các khoản chi trả khác cho người lao động", "196", form.phan3.cau16_nhanCong.ma196_chiTraKhac, form.phan3.cau16_nhanCong.gc_196]);
  rows.push(["TỔNG CỘNG NHANCONG", "NHANCONG", form.phan3.cau16_nhanCong.tongCong, form.phan3.cau16_nhanCong.gc_tong]);
  rows.push([]);

  // Khấu hao (Câu 17)
  rows.push(["17. Chi phí khấu hao TSCĐ dùng cho SXKD SP chính (Mã 225):", form.phan3.cau17_khauHao225]);
  rows.push([]);

  // Bảng Câu 18 (Dịch vụ mua ngoài & chi phí khác theo Mã IO)
  rows.push(["18. Chi phí dịch vụ mua ngoài và chi phí khác bằng tiền năm 2025 (theo Mã IO)", "", "", "", "", ""]);
  rows.push(["Mã I/O", "Tên loại chi phí / dịch vụ mua ngoài", "Trị giá (đồng) (1)", "Tỷ lệ nhập khẩu (%) (2)"]);
  form.phan3.cau18_muaNgoaiChiPhiKhac.forEach(item => {
    rows.push([item.maIO, item.moTa, item.c1_giaTri, `${item.c2_tyLeNK}%`]);
  });
  const tongC18 = form.phan3.cau18_muaNgoaiChiPhiKhac.reduce((s, r) => s + r.c1_giaTri, 0);
  rows.push(["TỔNG", "TỔNG CỘNG CHI PHÍ DỊCH VỤ MUA NGOÀI (CÂU 18)", tongC18, ""]);
  rows.push([]);

  // Nợ vay (Câu 19.4)
  rows.push(["19.4 Vay và nợ thuê tài chính ngắn hạn (Mã 320 CĐKT):", form.phan4.cau194_vayNganHan320]);
  rows.push(["19.4 Vay và nợ thuê tài chính dài hạn (Mã 338 CĐKT):", form.phan4.cau194_vayDaiHan338]);
  rows.push([]);

  // TỔNG HỢP CHỈ SỐ IO
  rows.push(["ĐỐI SOÁT CHỈ SỐ KINH TẾ TỔNG HỢP (IO 2026)", "", "", "", "", ""]);
  rows.push(["Chi phí trung gian (IC):", form.tongHopIO.ic]);
  rows.push(["Giá trị sản xuất (GO):", form.tongHopIO.go]);
  rows.push(["Tỷ lệ IC/GO:", `${(form.tongHopIO.tyLeIC_GO * 100).toFixed(2)}%`]);
  rows.push(["Đánh giá sơ bộ IC/GO:", form.tongHopIO.danhGiaIC]);

  const wsPhieu = XLSX.utils.aoa_to_sheet(rows);
  wsPhieu["!cols"] = [{ wch: 45 }, { wch: 25 }, { wch: 28 }, { wch: 28 }, { wch: 25 }, { wch: 25 }];
  XLSX.utils.book_append_sheet(wb, wsPhieu, "Phieu_01_IO_DN");

  // === SHEET 2: BIỂU GIẢI TRÌNH LỖI VÀ CẢNH BÁO (ĐỀ NGHỊ ĐTV GIẢI TRÌNH) ===
  const rowsLoi: any[][] = [];
  rowsLoi.push(["BIỂU GIẢI TRÌNH LỖI VÀ CẢNH BÁO PHIẾU 01/IO-DN", "", "", "", "", "", ""]);
  rowsLoi.push([`(Đề nghị Điều tra viên giải trình theo Quyết định 1124/QĐ-CTK)`, "", "", "", "", "", ""]);
  rowsLoi.push([`Doanh nghiệp: ${form.thongTinDinhDanh.tenDoanhNghiep} - MST: ${form.thongTinDinhDanh.maSoThue}`, "", "", "", "", "", ""]);
  rowsLoi.push([`Ngày kiểm tra: ${new Date().toLocaleDateString("vi-VN")}`, "", "", "", "", "", ""]);
  rowsLoi.push([]);
  rowsLoi.push(["STT", "Vị trí", "Phân loại", "Tên lỗi / Cảnh báo", "Quy định kiểm tra", "Số liệu thực tế", "Ý kiến giải trình của ĐTV"]);

  if (violations.length === 0) {
    rowsLoi.push(["-", "Toàn phiếu", "ĐẠT CHUẨN", "Phiếu hoàn toàn hợp lệ, không phát hiện lỗi logic nào", "Theo quy định IO 2026", "Hợp lệ", "Đã đối chiếu khớp đúng BCTC"]);
  } else {
    violations.forEach((v, idx) => {
      rowsLoi.push([
        idx + 1,
        v.cau,
        v.type === "error" ? "LỖI (Bắt buộc sửa)" : "CẢNH BÁO (Cần xác minh)",
        v.tenLoi,
        v.moTa,
        v.giaTriThucTe,
        "" // Cột để trống cho ĐTV điền ý kiến giải trình
      ]);
    });
  }

  const wsLoi = XLSX.utils.aoa_to_sheet(rowsLoi);
  wsLoi["!cols"] = [{ wch: 6 }, { wch: 14 }, { wch: 24 }, { wch: 38 }, { wch: 45 }, { wch: 25 }, { wch: 35 }];
  XLSX.utils.book_append_sheet(wb, wsLoi, "Bieu_Giai_Trinh_Loi");

  // Xuất file
  const safeName = form.thongTinDinhDanh.tenDoanhNghiep.replace(/[^a-zA-Z0-9_\u00C0-\u024F\u1EA0-\u1EF9]/g, "_").slice(0, 30);
  XLSX.writeFile(wb, `Phieu_01_IO_DN_${safeName}_${Date.now()}.xlsx`);
}

// Xuất bảng biểu tổng hợp lỗi cho danh sách nhiều doanh nghiệp
export function exportBatchAuditExcel(
  summaryList: Array<{
    stt: number;
    tenDN: string;
    mst: string;
    soLoi: number;
    soCanhBao: number;
    trangThai: string;
    dtt: number;
    ic: number;
    go: number;
    tyLeIC_GO: number;
    chiTietLoi: string;
  }>
) {
  const wb = XLSX.utils.book_new();
  const rows: any[][] = [];

  rows.push(["TỔNG HỢP KIỂM TRA LỖI DOANH NGHIỆP THEO HƯỚNG DẪN PHIẾU 01/IO-DN (IO 2026)", "", "", "", "", "", "", "", "", ""]);
  rows.push([`Ngày xuất báo cáo: ${new Date().toLocaleDateString("vi-VN")}`, "", "", "", "", "", "", "", "", ""]);
  rows.push([]);
  rows.push([
    "STT",
    "Mã số thuế",
    "Tên doanh nghiệp",
    "Trạng thái",
    "Số lỗi",
    "Số cảnh báo",
    "Doanh thu thuần (C5)",
    "IC (Chi phí TG)",
    "GO (Giá trị SX)",
    "Tỷ lệ IC/GO",
    "Chi tiết lỗi & Cảnh báo cần giải trình"
  ]);

  summaryList.forEach(item => {
    rows.push([
      item.stt,
      item.mst,
      item.tenDN,
      item.trangThai,
      item.soLoi,
      item.soCanhBao,
      item.dtt,
      item.ic,
      item.go,
      `${(item.tyLeIC_GO * 100).toFixed(1)}%`,
      item.chiTietLoi
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [
    { wch: 6 },
    { wch: 15 },
    { wch: 35 },
    { wch: 16 },
    { wch: 10 },
    { wch: 12 },
    { wch: 20 },
    { wch: 18 },
    { wch: 18 },
    { wch: 12 },
    { wch: 55 }
  ];
  XLSX.utils.book_append_sheet(wb, ws, "TongHop_KiemTra_IO2026");
  XLSX.writeFile(wb, `TongHop_KiemTra_Loi_DN_IO2026_${Date.now()}.xlsx`);
}

// Tạo đối tượng Workbook mẫu BCTC Excel chuẩn (B01-DN, B02-DN, F01-DN)
export function generateSampleBctcWorkbook(): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // SHEET 1: B02-DN - KẾT QUẢ HOẠT ĐỘNG KINH DOANH
  const b02Rows: any[][] = [
    ["CÔNG TY CỔ PHẦN CƠ KHÍ & CHẾ TẠO MÁY HƯNG YÊN"],
    ["Mã số thuế: 0900123456"],
    ["Địa chỉ: KCN Phố Nối A, Huyện Văn Lâm, Tỉnh Hưng Yên"],
    ["Mẫu số B 02 - DN (Ban hành theo Thông tư số 200/2014/TT-BTC)"],
    ["BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH NĂM 2025"],
    ["Đơn vị tính: Đồng"],
    [],
    ["Chỉ tiêu", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
    ["1. Doanh thu bán hàng và cung cấp dịch vụ", "01", "VI.25", 48500000000, 42000000000],
    ["2. Các khoản giảm trừ doanh thu", "02", "VI.26", 500000000, 200000000],
    ["3. Doanh thu thuần về bán hàng và CCDV (10 = 01 - 02)", "10", "VI.27", 48000000000, 41800000000],
    ["4. Giá vốn hàng bán", "11", "VI.28", 37500000000, 32600000000],
    ["5. Lợi nhuận gộp về bán hàng và CCDV (20 = 10 - 11)", "20", "", 10500000000, 9200000000],
    ["6. Doanh thu hoạt động tài chính", "21", "VI.29", 450000000, 380000000],
    ["7. Chi phí tài chính", "22", "VI.30", 1200000000, 950000000],
    ["- Trong đó: Chi phí lãi vay", "23", "", 1150000000, 910000000],
    ["8. Chi phí bán hàng", "25", "VI.31", 2400000000, 2100000000],
    ["9. Chi phí quản lý doanh nghiệp", "26", "VI.32", 2800000000, 2500000000],
    ["10. Lợi nhuận thuần từ HĐKD (30 = 20 + 21 - 22 - 25 - 26)", "30", "", 4550000000, 4030000000],
    ["11. Thu nhập khác", "31", "", 120000000, 80000000],
    ["12. Chi phí khác", "32", "", 50000000, 30000000],
    ["13. Lợi nhuận khác (40 = 31 - 32)", "40", "", 70000000, 50000000],
    ["14. Tổng lợi nhuận kế toán trước thuế (50 = 30 + 40)", "50", "", 4620000000, 4080000000],
    ["15. Chi phí thuế TNDN hiện hành", "51", "VI.33", 924000000, 816000000],
    ["16. Lợi nhuận sau thuế TNDN (60 = 50 - 51)", "60", "", 3696000000, 3264000000]
  ];
  const wsB02 = XLSX.utils.aoa_to_sheet(b02Rows);
  wsB02["!cols"] = [{ wch: 55 }, { wch: 10 }, { wch: 15 }, { wch: 18 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, wsB02, "B02-DN_KQKD");

  // SHEET 2: B01-DN - BẢNG CÂN ĐỐI KẾ TOÁN
  const b01Rows: any[][] = [
    ["CÔNG TY CỔ PHẦN CƠ KHÍ & CHẾ TẠO MÁY HƯNG YÊN"],
    ["Mã số thuế: 0900123456"],
    ["Địa chỉ: KCN Phố Nối A, Huyện Văn Lâm, Tỉnh Hưng Yên"],
    ["Mẫu số B 01 - DN (Ban hành theo Thông tư số 200/2014/TT-BTC)"],
    ["BẢNG CÂN ĐỐI KẾ TOÁN NĂM 2025"],
    ["Đơn vị tính: Đồng"],
    [],
    ["TÀI SẢN", "Mã số", "Thuyết minh", "Số cuối năm", "Số đầu năm"],
    ["A - TÀI SẢN NGẮN HẠN", "100", "", 32500000000, 28000000000],
    ["I. Tiền và các khoản tương đương tiền", "110", "V.01", 4500000000, 3800000000],
    ["IV. Hàng tồn kho", "140", "V.04", 12500000000, 10200000000],
    ["1. Hàng tồn kho", "141", "", 12500000000, 10200000000],
    ["B - TÀI SẢN DÀI HẠN", "200", "", 32500000000, 30000000000],
    ["II. Tài sản cố định", "220", "", 28500000000, 26000000000],
    ["1. Tài sản cố định hữu hình", "221", "V.08", 28500000000, 26000000000],
    ["- Nguyên giá", "222", "", 42000000000, 37500000000],
    ["- Giá trị hao mòn luỹ kế", "223", "", -13500000000, -11500000000],
    ["TỔNG CỘNG TÀI SẢN (270 = 100 + 200)", "270", "", 65000000000, 58000000000],
    [],
    ["NGUỒN VỐN", "Mã số", "Thuyết minh", "Số cuối năm", "Số đầu năm"],
    ["C - NỢ PHẢI TRẢ", "300", "", 28000000000, 24000000000],
    ["I. Nợ ngắn hạn", "310", "", 23500000000, 19000000000],
    ["4. Vay và nợ thuê tài chính ngắn hạn", "320", "V.15", 9500000000, 8200000000],
    ["II. Nợ dài hạn", "330", "", 4500000000, 5000000000],
    ["8. Vay và nợ thuê tài chính dài hạn", "338", "V.16", 4500000000, 5000000000],
    ["D - VỐN CHỦ SỞ HỮU", "400", "", 37000000000, 34000000000],
    ["I. Vốn chủ sở hữu", "410", "V.22", 37000000000, 34000000000],
    ["1. Vốn góp của chủ sở hữu", "411", "", 30000000000, 30000000000],
    ["11. Lợi nhuận sau thuế chưa phân phối", "421", "", 7000000000, 4000000000],
    ["TỔNG CỘNG NGUỒN VỐN (440 = 300 + 400)", "440", "", 65000000000, 58000000000]
  ];
  const wsB01 = XLSX.utils.aoa_to_sheet(b01Rows);
  wsB01["!cols"] = [{ wch: 50 }, { wch: 10 }, { wch: 15 }, { wch: 18 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, wsB01, "B01-DN_CDKT");

  // SHEET 3: F01-DN - BẢNG CÂN ĐỐI PHÁT SINH TÀI KHOẢN
  const f01Rows: any[][] = [
    ["CÔNG TY CỔ PHẦN CƠ KHÍ & CHẾ TẠO MÁY HƯNG YÊN"],
    ["Mã số thuế: 0900123456"],
    ["BẢNG CÂN ĐỐI SỐ PHÁT SINH TÀI KHOẢN NĂM 2025"],
    ["Đơn vị tính: Đồng"],
    [],
    ["Số hiệu TK", "Tên tài khoản", "Dư đầu kỳ Nợ", "Dư đầu kỳ Có", "Phát sinh Nợ", "Phát sinh Có", "Dư cuối kỳ Nợ", "Dư cuối kỳ Có"],
    ["111", "Tiền mặt", 500000000, 0, 12000000000, 11800000000, 700000000, 0],
    ["112", "Tiền gửi ngân hàng", 3300000000, 0, 55000000000, 54500000000, 3800000000, 0],
    ["152", "Nguyên liệu, vật liệu", 5800000000, 0, 24000000000, 22500000000, 7300000000, 0],
    ["154", "Chi phí SXKD dở dang", 1200000000, 0, 33100000000, 32850000000, 1450000000, 0],
    ["155", "Thành phẩm", 2100000000, 0, 32850000000, 32550000000, 2400000000, 0],
    ["157", "Hàng gửi đi bán", 450000000, 0, 8500000000, 8430000000, 520000000, 0],
    ["214", "Hao mòn TSCĐ", 0, 11500000000, 250000000, 2250000000, 0, 13500000000],
    ["334", "Phải trả người lao động", 0, 450000000, 6750000000, 6800000000, 0, 500000000],
    ["338", "Phải trả, phải nộp khác (BHXH)", 0, 120000000, 1530000000, 1550000000, 0, 140000000],
    ["511", "Doanh thu bán hàng và CCDV", 0, 0, 48000000000, 48000000000, 0, 0],
    ["621", "Chi phí NVL trực tiếp", 0, 0, 22500000000, 22500000000, 0, 0],
    ["622", "Chi phí nhân công trực tiếp", 0, 0, 6800000000, 6800000000, 0, 0],
    ["627", "Chi phí sản xuất chung", 0, 0, 3800000000, 3800000000, 0, 0],
    ["632", "Giá vốn hàng bán", 0, 0, 37500000000, 37500000000, 0, 0],
    ["641", "Chi phí bán hàng", 0, 0, 2400000000, 2400000000, 0, 0],
    ["642", "Chi phí quản lý doanh nghiệp", 0, 0, 2800000000, 2800000000, 0, 0]
  ];
  const wsF01 = XLSX.utils.aoa_to_sheet(f01Rows);
  wsF01["!cols"] = [{ wch: 12 }, { wch: 35 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, wsF01, "F01-DN_CDTK");

  return wb;
}

// Xuất tệp Excel BCTC Mẫu để người dùng tải về xem cấu trúc
export function exportSampleBctcExcelFile() {
  const wb = generateSampleBctcWorkbook();
  XLSX.writeFile(wb, "BCTC_Mau_TT200_HungYen_2025.xlsx");
}
