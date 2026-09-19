import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { 
  FileSpreadsheet, 
  Upload, 
  Download, 
  RefreshCw, 
  AlertTriangle, 
  Layers, 
  Shuffle, 
  Search, 
  Sliders,
  Database,
  Phone,
  ArrowRightLeft,
  RotateCcw
} from "lucide-react";

export interface HoRecord {
  sttGoc: any;
  hoTen: string;
  diaChi: string;
  nhanKhau: any;
  loaiHo: number;
  tenLoaiHo: string;
  ghiChu: string;
  trangThaiMau?: "CHÍNH THỨC" | "DỰ PHÒNG" | "MẤT MẪU (ĐÃ THAY)";
  sttMau?: number;
  sheetName?: string;
  tenDiaBan?: string;
  duPhongChoLoai?: number; // Loại nguồn thu mà hộ này sẵn sàng thay thế
  thayTheChoHo?: string;   // Tên/STT hộ đã thay thế
  raw?: any;
}

export interface DiaBanData {
  sheetName: string;
  tenDiaBan: string;
  tinh: string;
  xa: string;
  originalCount: number;
  loaiCounts: Record<number, number>;
  allHouseholds: HoRecord[];
  chinhThuc: HoRecord[];
  duPhong: HoRecord[];
  ketQua: HoRecord[];
  khongNongNghiep: boolean;
  soHoNongNghiep: number;
  tyLeNongNghiep: number;
}

export const LOAI_HO_META: Record<number, { ten: string; moTa: string; badgeColor: string; bgSoft: string }> = {
  1: {
    ten: "Nông, lâm thủy sản",
    moTa: "Nguồn thu lớn nhất từ ngành Nông lâm thủy sản",
    badgeColor: "bg-emerald-600 text-white",
    bgSoft: "bg-emerald-50 text-emerald-800 border-emerald-200"
  },
  2: {
    ten: "Công nghiệp xây dựng",
    moTa: "Nguồn thu lớn nhất từ ngành Công nghiệp xây dựng",
    badgeColor: "bg-blue-600 text-white",
    bgSoft: "bg-blue-50 text-blue-800 border-blue-200"
  },
  3: {
    ten: "Thương mại dịch vụ",
    moTa: "Nguồn thu lớn nhất từ ngành Thương mại dịch vụ",
    badgeColor: "bg-amber-600 text-white",
    bgSoft: "bg-amber-50 text-amber-800 border-amber-200"
  },
  4: {
    ten: "Nguồn thu khác / Lương",
    moTa: "Nguồn thu lớn nhất từ tiền lương, tiền công hoặc nguồn khác",
    badgeColor: "bg-purple-600 text-white",
    bgSoft: "bg-purple-50 text-purple-800 border-purple-200"
  }
};

// Thuật toán xáo trộn Fisher-Yates có thể lặp lại theo seed
function shuffleArray<T>(arr: T[], seedOffset: number = 0): T[] {
  const result = [...arr];
  let s = (12345 + seedOffset * 6789) % 2147483647;
  const rand = () => {
    s = (s * 16807) % 2147483647;
    return (s - 1) / 2147483646;
  };
  for (let i = result.length - 1; i > 0; i--) {
    const j = Math.floor(rand() * (i + 1));
    [result[i], result[j]] = [result[j], result[i]];
  }
  return result;
}

export const PhanBoHoXa: React.FC = () => {
  const [fileName, setFileName] = useState<string>("");
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [rawWorkbook, setRawWorkbook] = useState<XLSX.WorkBook | null>(null);

  // Thông tin hành chính cấp xã trích xuất tự động từ file
  const [communeMeta, setCommuneMeta] = useState<{ tinh: string; xa: string }>({
    tinh: "",
    xa: ""
  });

  // Cấu hình số hộ lấy mẫu (Mặc định 10 chính thức + 4 dự phòng)
  const [targetChinhThuc, setTargetChinhThuc] = useState<number>(10);
  const [targetDuPhong, setTargetDuPhong] = useState<number>(4);
  const [randomSeed, setRandomSeed] = useState<number>(42);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Dữ liệu từng địa bàn (mỗi sheet là 1 địa bàn)
  const [diaBanMap, setDiaBanMap] = useState<Record<string, DiaBanData>>({});
  const [activeSheetTab, setActiveSheetTab] = useState<string>("");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [filterLoai, setFilterLoai] = useState<string>("all");
  const [filterTrangThai, setFilterTrangThai] = useState<string>("all");
  const [replaceNotice, setReplaceNotice] = useState<{ message: string; type: "success" | "warning" } | null>(null);

  // Tổng hợp toàn xã
  const summaryXa = useMemo(() => {
    const sheets = Object.keys(diaBanMap);
    const tongLoai: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0 };
    let tongHo = 0;
    let tongChinhThuc = 0;
    let tongDuPhong = 0;

    const diaBanKhongNN: string[] = [];
    const diaBanCoNN: string[] = [];

    sheets.forEach((sheetName) => {
      const db = diaBanMap[sheetName];
      const hoNN = db.loaiCounts[1] || 0;
      if (hoNN === 0) {
        diaBanKhongNN.push(sheetName);
      } else {
        diaBanCoNN.push(sheetName);
      }
      tongHo += db.originalCount;
      tongChinhThuc += db.chinhThuc.length;
      tongDuPhong += db.duPhong.length;
      for (let l = 1; l <= 4; l++) {
        tongLoai[l] += db.loaiCounts[l] || 0;
      }
    });

    // Thuật toán Phần Dư Lớn Nhất (Largest Remainder Method / Hamilton Method)
    const totalMauChinhThucXa = sheets.length * targetChinhThuc;
    const mauChiaLoai: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0 };
    const remainders: { loai: number; rem: number }[] = [];

    if (tongHo > 0) {
      let sumAllocated = 0;
      for (let l = 1; l <= 4; l++) {
        const exactVal = (tongLoai[l] / tongHo) * totalMauChinhThucXa;
        const intVal = Math.floor(exactVal);
        mauChiaLoai[l] = intVal;
        sumAllocated += intVal;
        remainders.push({ loai: l, rem: exactVal - intVal });
      }

      const thieu = totalMauChinhThucXa - sumAllocated;
      remainders.sort((a, b) => b.rem - a.rem);
      for (let i = 0; i < thieu && i < remainders.length; i++) {
        mauChiaLoai[remainders[i].loai] += 1;
      }
    }

    return {
      numSheets: sheets.length,
      tongHo,
      tongLoai,
      tongChinhThuc,
      tongDuPhong,
      totalMauChinhThucXa,
      mauChiaLoai,
      diaBanKhongNN,
      diaBanCoNN,
      countKhongNN: diaBanKhongNN.length,
      countCoNN: diaBanCoNN.length
    };
  }, [diaBanMap, targetChinhThuc]);

  // Phân tích workbook Excel
  const parseWorkbook = (wb: XLSX.WorkBook, seed: number) => {
    setIsProcessing(true);
    setErrorMessage("");

    try {
      const newDiaBanMap: Record<string, DiaBanData> = {};
      const allSheetNames = wb.SheetNames;
      let detectedTinh = "";
      let detectedXa = "";

      for (const sheetName of allSheetNames) {
        const ws = wb.Sheets[sheetName];
        if (!ws) continue;

        const rawGrid: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (!rawGrid || rawGrid.length === 0) continue;

        // Trích xuất metadata hành chính từ 8 dòng đầu
        let sheetTinh = "";
        let sheetXa = "";
        let sheetDiaBanName = sheetName;

        for (let r = 0; r < Math.min(rawGrid.length, 8); r++) {
          const row = rawGrid[r];
          for (let c = 0; c < row.length; c++) {
            const val = String(row[c] || "").trim();
            if (!val) continue;

            const matchTinh = val.match(/tỉnh\s*:\s*([^\n\r]+)/i);
            if (matchTinh && matchTinh[1].trim()) {
              sheetTinh = matchTinh[1].trim();
              if (!detectedTinh) detectedTinh = sheetTinh;
            }

            const matchXa = val.match(/(?:xã\/phường|xã|phường|thị trấn)\s*:\s*([^\n\r]+)/i);
            if (matchXa && matchXa[1].trim()) {
              sheetXa = matchXa[1].trim();
              if (!detectedXa) detectedXa = sheetXa;
            }

            const matchDb = val.match(/tên địa bàn\s*:\s*([^\n\r]+)/i);
            if (matchDb && matchDb[1].trim()) {
              sheetDiaBanName = matchDb[1].trim();
            }
          }
        }

        // Tìm dòng tiêu đề (Header Row)
        let headerRowIndex = -1;
        let sttCol = 0;
        let hoTenCol = 1;
        let diaChiCol = 2;
        let nhanKhauCol = 3;
        let loaiCol = 4;
        let ghiChuCol = 5;

        for (let r = 0; r < Math.min(rawGrid.length, 25); r++) {
          const row = rawGrid[r];
          let foundLoai = false;
          let foundHoTen = false;

          for (let c = 0; c < row.length; c++) {
            const cellVal = String(row[c] || "").toLowerCase().trim();
            if (
              cellVal.includes("nguồn thu nhập lớn nhất") ||
              cellVal.includes("thu nhập lớn nhất") ||
              cellVal.includes("loại hộ")
            ) {
              loaiCol = c;
              foundLoai = true;
            }
            if (cellVal.includes("họ và tên") || cellVal.includes("chủ hộ")) {
              hoTenCol = c;
              foundHoTen = true;
            }
            if (cellVal === "stt" || cellVal.startsWith("stt ")) {
              sttCol = c;
            }
            if (cellVal.includes("địa chỉ") || cellVal.includes("thôn") || cellVal.includes("xóm") || cellVal.includes("tổ dân phố")) {
              diaChiCol = c;
            }
            if (cellVal.includes("nhân khẩu")) {
              nhanKhauCol = c;
            }
            if (cellVal.includes("ghi chú") || cellVal.includes("điện thoại") || cellVal.includes("sđt")) {
              ghiChuCol = c;
            }
          }

          if (foundLoai || foundHoTen) {
            headerRowIndex = r;
            break;
          }
        }

        // XÁC ĐỊNH VÀ BỎ QUA DÒNG KÝ HIỆU (A, B, C, 1, 2, 3...) DƯỚI TIÊU ĐỀ
        // Đúng theo ảnh người dùng cung cấp:
        // Dòng header: STT, Họ và tên chủ hộ, Địa chỉ..., Số nhân khẩu..., Nguồn thu nhập..., Ghi chú
        // Dòng ký hiệu dưới tiêu đề: A, B, C, 1, 2, 3
        // Dòng dữ liệu bắt đầu sau dòng ký hiệu
        let startDataRow = headerRowIndex !== -1 ? headerRowIndex + 1 : 4;
        
        // Kiểm tra xem dòng ngay sau header có phải là dòng ký hiệu hay không
        if (startDataRow < rawGrid.length) {
          const checkRow = rawGrid[startDataRow];
          const cA = String(checkRow[sttCol] || "").trim().toUpperCase();
          const cB = String(checkRow[hoTenCol] || "").trim().toUpperCase();
          const cC = String(checkRow[diaChiCol] || "").trim().toUpperCase();
          
          if (cA === "A" || cB === "B" || cC === "C" || cA === "STT") {
            // Đây chính xác là dòng ký hiệu, bỏ qua và bắt đầu đọc từ dòng kế tiếp
            startDataRow = startDataRow + 1;
          }
        }

        const validHouseholds: HoRecord[] = [];
        const loaiCounts: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0 };

        for (let r = startDataRow; r < rawGrid.length; r++) {
          const row = rawGrid[r];
          if (!row || row.length === 0) continue;

          const rawStt = String(row[sttCol] || "").trim();
          const rawHoTen = String(row[hoTenCol] || "").trim();
          const rawDiaChi = String(row[diaChiCol] || "").trim();
          const rawNhanKhau = row[nhanKhauCol] !== undefined && row[nhanKhauCol] !== "" ? row[nhanKhauCol] : 1;
          let rawLoai = row[loaiCol];
          const rawGhiChu = String(row[ghiChuCol] || "").trim();

          // Kiểm tra loại trừ: Nếu dòng vẫn là dòng ký hiệu (A, B, C) hoặc dòng header lặp lại
          if (rawStt.toUpperCase() === "A" || rawHoTen.toUpperCase() === "B") continue;
          if (rawHoTen.toLowerCase().includes("họ và tên") || rawStt.toLowerCase() === "stt") continue;
          
          // Bỏ qua dòng tổng cộng hoặc người ký cuối bảng
          if (
            rawHoTen.toLowerCase().includes("tổng") || 
            rawHoTen.toLowerCase().includes("cộng") ||
            rawHoTen.toLowerCase().includes("người lập") ||
            rawHoTen.toLowerCase().includes("trưởng thôn")
          ) {
            continue;
          }

          // Dò tìm giá trị loại hộ (1, 2, 3, 4)
          if (rawLoai === undefined || rawLoai === null || String(rawLoai).trim() === "") {
            for (let c = 0; c < Math.min(row.length, 8); c++) {
              const numVal = parseInt(String(row[c]).trim(), 10);
              if ([1, 2, 3, 4].includes(numVal) && c !== sttCol && c !== hoTenCol) {
                rawLoai = numVal;
                break;
              }
            }
          }

          const loaiNum = parseInt(String(rawLoai).trim(), 10);
          if (![1, 2, 3, 4].includes(loaiNum)) continue; // Chỉ giữ các dòng có Loại hộ 1..4 hợp lệ

          // Tên chủ hộ hợp lệ
          if (!rawHoTen || rawHoTen.length < 2) continue;

          const stt = rawStt ? (isNaN(Number(rawStt)) ? rawStt : Number(rawStt)) : validHouseholds.length + 1;

          const hoRecord: HoRecord = {
            sttGoc: stt,
            hoTen: rawHoTen,
            diaChi: rawDiaChi,
            nhanKhau: rawNhanKhau,
            loaiHo: loaiNum,
            tenLoaiHo: LOAI_HO_META[loaiNum]?.ten || "Khác",
            ghiChu: rawGhiChu,
            sheetName,
            tenDiaBan: sheetDiaBanName,
            raw: row
          };

          validHouseholds.push(hoRecord);
          loaiCounts[loaiNum] = (loaiCounts[loaiNum] || 0) + 1;
        }

        if (validHouseholds.length > 0) {
          const soHoNongNghiep = loaiCounts[1] || 0;
          const tyLeNongNghiep = validHouseholds.length > 0 ? (soHoNongNghiep / validHouseholds.length) * 100 : 0;
          newDiaBanMap[sheetName] = {
            sheetName,
            tenDiaBan: sheetDiaBanName,
            tinh: sheetTinh || detectedTinh,
            xa: sheetXa || detectedXa,
            originalCount: validHouseholds.length,
            loaiCounts,
            allHouseholds: validHouseholds,
            chinhThuc: [],
            duPhong: [],
            ketQua: [],
            khongNongNghiep: soHoNongNghiep === 0,
            soHoNongNghiep,
            tyLeNongNghiep
          };
        }
      }

      setCommuneMeta({
        tinh: detectedTinh || "Chưa xác định",
        xa: detectedXa || "Chưa xác định"
      });

      const validSheets = Object.keys(newDiaBanMap);
      if (validSheets.length === 0) {
        setErrorMessage("Không tìm thấy dữ liệu bảng kê hợp lệ (chứa cột Loại hộ 1-4). Vui lòng kiểm tra lại cấu trúc file Excel.");
        setIsProcessing(false);
        return;
      }

      // Tính tổng toàn xã của từng loại
      const tongXaLoai: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0 };
      let totalHoXa = 0;
      validSheets.forEach((sn) => {
        totalHoXa += newDiaBanMap[sn].originalCount;
        for (let l = 1; l <= 4; l++) {
          tongXaLoai[l] += newDiaBanMap[sn].loaiCounts[l] || 0;
        }
      });

      // BỐC MẪU CHO TỪNG ĐỊA BÀN
      validSheets.forEach((sn, sheetIdx) => {
        const db = newDiaBanMap[sn];
        const dfDb = [...db.allHouseholds];

        let hoChinhThucDb: HoRecord[] = [];
        const seedShift = seed + sheetIdx * 23;
        const isKhongNN = db.khongNongNghiep;

        // 1. Phân bổ chỉ tiêu hộ chính thức cho địa bàn:
        // Đếm số lượng hộ từng loại có trong địa bàn
        const soLuongGoc: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0 };
        for (let l = 1; l <= 4; l++) {
          soLuongGoc[l] = dfDb.filter((h) => h.loaiHo === l).length;
        }

        // Xác định các loại có thể lấy chính thức:
        // - Địa bàn không có nông nghiệp: chỉ lấy Loại 2, 3, 4
        // - Địa bàn có nông nghiệp: lấy Loại 1, 2, 3, 4
        const loaiUngVien = (isKhongNN ? [2, 3, 4] : [1, 2, 3, 4]).filter(
          (l) => soLuongGoc[l] > 0
        );

        // QUAN TRỌNG: Để bảo đảm hộ dự phòng LUÔN có đủ đại diện từng loại,
        // nếu một loại có từ 2 hộ trở lên, ta chừa lại ít nhất 1 hộ cho dự phòng!
        const maxChinhThucChoLoai: Record<number, number> = {};
        loaiUngVien.forEach((l) => {
          maxChinhThucChoLoai[l] = soLuongGoc[l] > 1 ? soLuongGoc[l] - 1 : soLuongGoc[l];
        });

        // Tính phân bổ tỷ lệ theo phương pháp Phần Dư Lớn Nhất (Hamilton)
        const tongHoUngVien = loaiUngVien.reduce((sum, l) => sum + soLuongGoc[l], 0);
        const chiTieuChinhThuc: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0 };
        const phanDuCT: { loai: number; du: number; tong: number }[] = [];

        if (tongHoUngVien > 0) {
          loaiUngVien.forEach((l) => {
            // Tỷ trọng nội bộ của loại trong các loại ứng viên của địa bàn
            const tyTrong = soLuongGoc[l] / tongHoUngVien;
            const chiTieuThucTe = tyTrong * targetChinhThuc;
            const phanNguyen = Math.min(Math.floor(chiTieuThucTe), maxChinhThucChoLoai[l]);
            chiTieuChinhThuc[l] = phanNguyen;
            phanDuCT.push({
              loai: l,
              du: chiTieuThucTe - phanNguyen,
              tong: soLuongGoc[l]
            });
          });

          // Phân bổ phần còn thiếu cho các loại có phần dư lớn nhất
          let thieuCT = targetChinhThuc - Object.values(chiTieuChinhThuc).reduce((a, b) => a + b, 0);

          // Sắp xếp theo phần dư giảm dần, nếu bằng nhau thì ưu tiên loại có số hộ nhiều hơn
          phanDuCT.sort((a, b) => b.du - a.du || b.tong - a.tong);

          for (const item of phanDuCT) {
            if (thieuCT <= 0) break;
            if (chiTieuChinhThuc[item.loai] < maxChinhThucChoLoai[item.loai]) {
              chiTieuChinhThuc[item.loai] += 1;
              thieuCT -= 1;
            }
          }

          // Nếu vẫn còn thiếu (do giới hạn maxChinhThuc), mở rộng lấy nốt các hộ còn lại
          if (thieuCT > 0) {
            for (const item of phanDuCT) {
              if (thieuCT <= 0) break;
              if (chiTieuChinhThuc[item.loai] < soLuongGoc[item.loai]) {
                chiTieuChinhThuc[item.loai] += 1;
                thieuCT -= 1;
              }
            }
          }
        }

        // Bốc ngẫu nhiên hộ chính thức theo đúng chỉ tiêu đã tính
        loaiUngVien.forEach((l) => {
          const soCanLay = chiTieuChinhThuc[l] || 0;
          if (soCanLay > 0) {
            const hoLoaiL = dfDb.filter((h) => h.loaiHo === l);
            const shuffled = shuffleArray(hoLoaiL, seedShift + l * 37);
            hoChinhThucDb.push(...shuffled.slice(0, soCanLay));
          }
        });

        // Gán nhãn CHÍNH THỨC
        hoChinhThucDb = hoChinhThucDb.map((h, idx) => ({
          ...h,
          trangThaiMau: "CHÍNH THỨC" as const,
          sttMau: idx + 1
        }));

        // 3. Chọn hộ dự phòng (targetDuPhong = 4):
        // NGUYÊN TẮC THỐNG KÊ QUAN TRỌNG:
        // Hộ dự phòng dùng để THAY THẾ CHO HỘ BỊ MẤT MẪU (vắng nhà, chuyển đi, từ chối điều tra...).
        // Khi mất mẫu loại nào thì BẮT BUỘC phải thay bằng hộ dự phòng CÙNG LOẠI NGUỒN THU ĐÓ trong cùng địa bàn.
        // Do đó:
        // - Nếu địa bàn chọn toàn bộ 10 hộ là Thương mại (Loại 3) -> Toàn bộ 4 hộ dự phòng BẮT BUỘC là Thương mại (Loại 3).
        // - Nếu địa bàn có nhiều loại trong mẫu chính thức -> Hộ dự phòng phân bổ theo cơ cấu các loại đã chọn chính thức,
        //   đảm bảo loại nào có mẫu chính thức cũng có sẵn hộ dự phòng cùng loại để sẵn sàng thay thế.
        const sttChinhThuc = new Set(hoChinhThucDb.map((h) => h.sttGoc));
        const hoChuaChon = dfDb.filter((h) => !sttChinhThuc.has(h.sttGoc));

        // Đếm số hộ còn lại theo từng loại trong hoChuaChon
        const soLuongConLai: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0 };
        for (let l = 1; l <= 4; l++) {
          soLuongConLai[l] = hoChuaChon.filter((h) => h.loaiHo === l).length;
        }

        // Đếm số hộ chính thức đã chọn theo từng loại
        const soLuongCT: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0 };
        hoChinhThucDb.forEach((h) => {
          soLuongCT[h.loaiHo] = (soLuongCT[h.loaiHo] || 0) + 1;
        });

        // Các loại CÓ MẶT trong mẫu chính thức của địa bàn và vẫn còn hộ chưa chọn
        const loaiChinhThucCoMat = [1, 2, 3, 4].filter((l) => (soLuongCT[l] || 0) > 0);
        const chiTieuDuPhong: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0 };

        if (loaiChinhThucCoMat.length > 0) {
          // Bước 1: Cấp ít nhất 1 hộ dự phòng cho mỗi loại có trong mẫu chính thức (nếu còn hộ trong bảng kê)
          loaiChinhThucCoMat.forEach((l) => {
            if (soLuongConLai[l] > 0) {
              chiTieuDuPhong[l] = 1;
            }
          });

          // Bước 2: Số hộ dự phòng còn thiếu được phân bổ tiếp theo tỷ trọng số hộ chính thức
          // (Loại nào có nhiều hộ chính thức hơn thì nguy cơ mất mẫu cao hơn, cần nhiều dự phòng hơn)
          let thieuDP = targetDuPhong - Object.values(chiTieuDuPhong).reduce((a, b) => a + b, 0);

          // Sắp xếp các loại chính thức theo số hộ chính thức giảm dần, sau đó theo số hộ còn lại
          const loaiUuTien = [...loaiChinhThucCoMat].sort(
            (a, b) => (soLuongCT[b] || 0) - (soLuongCT[a] || 0) || (soLuongConLai[b] || 0) - (soLuongConLai[a] || 0)
          );

          while (thieuDP > 0) {
            let coThem = false;
            for (const l of loaiUuTien) {
              if (thieuDP <= 0) break;
              if (chiTieuDuPhong[l] < soLuongConLai[l]) {
                chiTieuDuPhong[l] += 1;
                thieuDP -= 1;
                coThem = true;
              }
            }
            if (!coThem) break; // Đã hết hộ của các loại chính thức
          }

          // Bước 3: Trường hợp các loại chính thức đã hết sạch hộ trong bảng kê, mới lấy loại khác còn lại
          let conThieu = targetDuPhong - Object.values(chiTieuDuPhong).reduce((a, b) => a + b, 0);
          if (conThieu > 0) {
            const loaiKhac = [1, 2, 3, 4].filter((l) => soLuongConLai[l] > chiTieuDuPhong[l]);
            for (const l of loaiKhac) {
              if (conThieu <= 0) break;
              const coTheLay = soLuongConLai[l] - chiTieuDuPhong[l];
              const lay = Math.min(conThieu, coTheLay);
              chiTieuDuPhong[l] += lay;
              conThieu -= lay;
            }
          }
        } else {
          // Phòng trường hợp đặc biệt không có mẫu chính thức
          const loaiCo = [1, 2, 3, 4].filter((l) => soLuongConLai[l] > 0);
          loaiCo.forEach((l) => {
            if (targetDuPhong > Object.values(chiTieuDuPhong).reduce((a, b) => a + b, 0)) {
              chiTieuDuPhong[l] = 1;
            }
          });
        }

        // Bốc ngẫu nhiên các hộ dự phòng theo đúng chỉ tiêu đã tính
        let hoDuPhongDb: HoRecord[] = [];
        [1, 2, 3, 4].forEach((l) => {
          const soCanLay = chiTieuDuPhong[l] || 0;
          if (soCanLay > 0) {
            const hoLoaiL = hoChuaChon.filter((h) => h.loaiHo === l);
            const shuffled = shuffleArray(hoLoaiL, seedShift + l * 79 + 53);
            const layRa = shuffled.slice(0, soCanLay).map((h) => ({
              ...h,
              duPhongChoLoai: l
            }));
            hoDuPhongDb.push(...layRa);
          }
        });

        // Gán nhãn DỰ PHÒNG
        hoDuPhongDb = hoDuPhongDb.map((h, idx) => ({
          ...h,
          trangThaiMau: "DỰ PHÒNG" as const,
          sttMau: hoChinhThucDb.length + idx + 1
        }));

        // Gộp lại thành bảng kết quả địa bàn
        db.chinhThuc = hoChinhThucDb;
        db.duPhong = hoDuPhongDb;
        db.ketQua = [...hoChinhThucDb, ...hoDuPhongDb];
      });

      setDiaBanMap(newDiaBanMap);
      if (!activeSheetTab || !newDiaBanMap[activeSheetTab]) {
        setActiveSheetTab(validSheets[0]);
      }
    } catch (err: any) {
      console.error(err);
      setErrorMessage(`Lỗi phân tích file Excel: ${err?.message || "Không xác định"}`);
    } finally {
      setIsProcessing(false);
    }
  };

  // Thay thế hộ bị mất mẫu bằng hộ dự phòng CÙNG LOẠI NGUỒN THU trong cùng địa bàn
  const handleReplaceHousehold = (sheetName: string, lostHouseholdSttGoc: any) => {
    setDiaBanMap((prev) => {
      const db = prev[sheetName];
      if (!db) return prev;

      const lostHo = db.chinhThuc.find((h) => h.sttGoc === lostHouseholdSttGoc);
      if (!lostHo) return prev;

      // Tìm hộ dự phòng CÙNG LOẠI NGUỒN THU trong cùng địa bàn
      const availableReserveIdx = db.duPhong.findIndex(
        (h) => h.loaiHo === lostHo.loaiHo
      );

      if (availableReserveIdx === -1) {
        setReplaceNotice({
          message: `Không còn hộ dự phòng cùng Loại ${lostHo.loaiHo} (${lostHo.tenLoaiHo}) trong địa bàn "${db.tenDiaBan || sheetName}" để thay thế! Nguyên tắc thống kê bắt buộc phải thay bằng hộ cùng loại nguồn thu để tránh méo mó cơ cấu.`,
          type: "warning"
        });
        return prev;
      }

      const replacementHo = db.duPhong[availableReserveIdx];

      // Đưa hộ dự phòng lên làm chính thức thay thế
      const updatedReplacement: HoRecord = {
        ...replacementHo,
        trangThaiMau: "CHÍNH THỨC",
        sttMau: lostHo.sttMau,
        thayTheChoHo: `${lostHo.hoTen} (STT gốc ${lostHo.sttGoc})`,
        ghiChu: replacementHo.ghiChu 
          ? `${replacementHo.ghiChu} | Thay cho: ${lostHo.hoTen}`
          : `Thay cho hộ: ${lostHo.hoTen} (STT gốc ${lostHo.sttGoc})`
      };

      // Đánh dấu hộ cũ bị mất mẫu (đã thay thế)
      const updatedLostHo: HoRecord = {
        ...lostHo,
        trangThaiMau: "MẤT MẪU (ĐÃ THAY)",
        thayTheChoHo: `Đã thay bằng: ${replacementHo.hoTen} (STT gốc ${replacementHo.sttGoc})`,
        ghiChu: lostHo.ghiChu 
          ? `${lostHo.ghiChu} | Mất mẫu (đã thay bằng ${replacementHo.hoTen})`
          : `Mất mẫu (đã thay bằng ${replacementHo.hoTen})`
      };

      // Cập nhật danh sách chính thức
      const newChinhThuc = db.chinhThuc.map((h) =>
        h.sttGoc === lostHouseholdSttGoc ? updatedReplacement : h
      );

      // Cập nhật danh sách dự phòng (bỏ hộ vừa được chọn)
      const newDuPhong = db.duPhong.filter((_, idx) => idx !== availableReserveIdx);

      // Cập nhật toàn bộ danh sách địa bàn
      const existingLost = (db.ketQua || []).filter((h) => h.trangThaiMau === "MẤT MẪU (ĐÃ THAY)");
      const newKetQua = [...newChinhThuc, ...newDuPhong, ...existingLost, updatedLostHo];

      setReplaceNotice({
        message: `Đã thay thế hộ mất mẫu "${lostHo.hoTen}" bằng hộ dự phòng "${replacementHo.hoTen}" (cùng Loại ${lostHo.loaiHo}: ${lostHo.tenLoaiHo}) tại địa bàn "${db.tenDiaBan || sheetName}". Cơ cấu và số lượng mẫu được bảo toàn 100%!`,
        type: "success"
      });

      return {
        ...prev,
        [sheetName]: {
          ...db,
          chinhThuc: newChinhThuc,
          duPhong: newDuPhong,
          ketQua: newKetQua
        }
      };
    });
  };

  // Nạp file Excel từ máy tính
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setIsProcessing(true);
    setErrorMessage("");

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        setRawWorkbook(wb);
        parseWorkbook(wb, randomSeed);
      } catch (err: any) {
        setErrorMessage(`Lỗi đọc file: ${err?.message || "File không đúng định dạng Excel"}`);
        setIsProcessing(false);
      }
    };
    reader.onerror = () => {
      setErrorMessage("Không thể đọc file đã chọn.");
      setIsProcessing(false);
    };
    reader.readAsArrayBuffer(file);
  };

  // Tái bốc mẫu ngẫu nhiên (đổi seed)
  const handleReshuffle = () => {
    if (!rawWorkbook) return;
    const newSeed = Math.floor(Math.random() * 100000);
    setRandomSeed(newSeed);
    parseWorkbook(rawWorkbook, newSeed);
  };

  // Tạo dữ liệu demo theo đúng mẫu phiếu trong ảnh (Xã Phường Mỹ Hào - Hưng Yên)
  const handleLoadDemoData = () => {
    setIsProcessing(true);
    setErrorMessage("");

    try {
      const demoWb = XLSX.utils.book_new();
      const hoMau = [
        { hoTen: "Phạm Văn Lộc", diaChi: "Tổ dân phố Nhuận Xá", nhanKhau: 4, loai: 3, sdt: "0866858232" },
        { hoTen: "Phạm Văn Cải", diaChi: "Tổ dân phố Nhuận Xá", nhanKhau: 6, loai: 2, sdt: "0333432404" },
        { hoTen: "Phạm Thị Hương Quỳnh", diaChi: "Tổ dân phố Nhuận Xá", nhanKhau: 6, loai: 2, sdt: "0968398861" },
        { hoTen: "Vũ Văn Định", diaChi: "TDP Tiên Xá 1", nhanKhau: 3, loai: 1, sdt: "0982345671" },
        { hoTen: "Trần Thị Mai", diaChi: "TDP Tiên Xá 2", nhanKhau: 4, loai: 3, sdt: "0971234589" },
        { hoTen: "Nguyễn Văn Hùng", diaChi: "TDP Tiên Xá 3", nhanKhau: 5, loai: 4, sdt: "0912349876" },
        { hoTen: "Đỗ Văn Thành", diaChi: "Tổ dân phố Nhuận Xá", nhanKhau: 4, loai: 1, sdt: "0967891234" },
        { hoTen: "Lê Thị Bích", diaChi: "TDP Tiên Xá 1", nhanKhau: 2, loai: 3, sdt: "0904567812" },
        { hoTen: "Hoàng Văn Tuấn", diaChi: "TDP Tiên Xá 2", nhanKhau: 5, loai: 2, sdt: "0934567890" },
        { hoTen: "Bùi Thị Nguyệt", diaChi: "TDP Tiên Xá 3", nhanKhau: 3, loai: 4, sdt: "0945678123" }
      ];

      // Tạo mẫu 12 địa bàn tượng trưng cho một xã có nhiều địa bàn
      const diaBanList = [
        "TDP Tiên Xá 1", "TDP Tiên Xá 2", "TDP Tiên Xá 3", "TDP Nhuận Xá 1", "TDP Nhuận Xá 2",
        "TDP Phố Nối 1", "TDP Phố Nối 2", "TDP Bến Xanh", "TDP Lương Bằng", "TDP Minh Đức",
        "TDP Hảo Xuyên", "TDP Thụy Trang"
      ];

      diaBanList.forEach((dbName, idx) => {
        const sheetName = `ĐB_${idx < 9 ? "0" + (idx + 1) : (idx + 1)}`;
        const rows: any[][] = [
          ["BẢNG KÊ HỘ", "", "", "", "", ""],
          ["ĐIỀU TRA THU NHẬP BÌNH QUÂN ĐẦU NGƯỜI CẤP XÃ", "", "", "", "", ""],
          ["Tỉnh: HƯNG YÊN", "", "", "Xã/phường: Phường Mỹ Hào", "", ""],
          [`Tên địa bàn: ${dbName}`, "", "", "Thành thị/Nông thôn (TTNT)", "", ""],
          ["Người rà soát bảng kê: ________________________", "", "", "Số điện thoại: ________________________", "", ""],
          [],
          [
            "STT", 
            "Họ và tên chủ hộ", 
            "Địa chỉ của hộ", 
            "Số nhân khẩu thực tế thường trú của hộ khi lập bảng kê", 
            "Nguồn thu nhập lớn nhất của hộ thuộc ngành nào (1: Nông lâm thủy sản; 2: Công nghiệp xây dựng; 3: Thương mại dịch vụ; 4: Nguồn khác)", 
            "Ghi chú"
          ],
          // DÒNG KÝ HIỆU DƯỚI TIÊU ĐỀ THEO ĐÚNG ẢNH MẪU:
          ["A", "B", "C", "1", "2", "3"]
        ];

        // Tạo 40 - 65 hộ trong mỗi địa bàn
        const numH = 45 + ((idx * 5) % 20);
        // Địa bàn 5, 6, 7, 8 tượng trưng cho các khu phố chợ/công nghiệp KHÔNG CÓ hộ nông nghiệp (0 hộ)
        const isDbPhiNongNghiep = idx >= 5 && idx <= 8;

        for (let h = 1; h <= numH; h++) {
          const sampleItem = hoMau[(h + idx) % hoMau.length];
          let loai = ((h + idx * 2) % 4) + 1;
          if (idx === 6) {
            // Địa bàn phố thương mại buôn bán: 100% hộ là Thương mại dịch vụ (Loại 3)
            loai = 3;
          } else if (isDbPhiNongNghiep && loai === 1) {
            // Chuyển loại 1 thành loại 2 (CNXD) hoặc loại 3 (Thương mại dịch vụ)
            loai = (h % 2 === 0) ? 2 : 3;
          }
          rows.push([
            h,
            `${sampleItem.hoTen} (${idx + 1})`,
            sampleItem.diaChi,
            sampleItem.nhanKhau,
            loai,
            sampleItem.sdt
          ]);
        }

        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(demoWb, ws, sheetName);
      });

      setFileName("Bang_Ke_Ho_Phuong_My_Hao_12_DiaBan.xlsx");
      setRawWorkbook(demoWb);
      parseWorkbook(demoWb, randomSeed);
    } catch (err: any) {
      setErrorMessage(`Lỗi tạo dữ liệu mẫu: ${err?.message}`);
      setIsProcessing(false);
    }
  };

  // Xuất file Excel kết quả theo đúng cấu trúc phiếu điều tra trong ảnh
  const handleExportExcel = () => {
    const sheets = Object.keys(diaBanMap);
    if (sheets.length === 0) return;

    try {
      const wb = XLSX.utils.book_new();
      const tinhName = communeMeta.tinh || "HƯNG YÊN";
      const xaName = communeMeta.xa || "CẤP XÃ";

      // Sheet 1: Tổng hợp toàn xã
      const summaryRows: any[][] = [
        ["BÁO CÁO PHÂN BỔ MẪU HỘ ĐIỀU TRA THU NHẬP - TOÀN XÃ"],
        [`Tỉnh: ${tinhName} | Xã/phường: ${xaName}`],
        [`File bảng kê gốc: ${fileName || "Chưa đặt tên"}`],
        [`Thời gian xuất file: ${new Date().toLocaleDateString("vi-VN")} ${new Date().toLocaleTimeString("vi-VN")}`],
        [`Quy mô tổng hợp: ${summaryXa.numSheets} địa bàn | ${summaryXa.tongHo} tổng hộ | ${summaryXa.tongChinhThuc} hộ chính thức | ${summaryXa.tongDuPhong} hộ dự phòng`],
        [],
        ["1. CƠ CẤU NGUỒN THU NHẬP VÀ PHÂN BỔ HẠN NGẠCH TOÀN XÃ"],
        ["Mã loại", "Tên nhóm ngành nguồn thu lớn nhất", "Tổng số hộ trong xã", "Tỷ trọng cơ cấu (%)", "Chỉ tiêu phân bổ mẫu chính thức toàn xã"],
        [1, LOAI_HO_META[1].ten, summaryXa.tongLoai[1], summaryXa.tongHo ? (summaryXa.tongLoai[1] / summaryXa.tongHo) : 0, summaryXa.mauChiaLoai[1]],
        [2, LOAI_HO_META[2].ten, summaryXa.tongLoai[2], summaryXa.tongHo ? (summaryXa.tongLoai[2] / summaryXa.tongHo) : 0, summaryXa.mauChiaLoai[2]],
        [3, LOAI_HO_META[3].ten, summaryXa.tongLoai[3], summaryXa.tongHo ? (summaryXa.tongLoai[3] / summaryXa.tongHo) : 0, summaryXa.mauChiaLoai[3]],
        [4, LOAI_HO_META[4].ten, summaryXa.tongLoai[4], summaryXa.tongHo ? (summaryXa.tongLoai[4] / summaryXa.tongHo) : 0, summaryXa.mauChiaLoai[4]],
        ["TỔNG", "TOÀN XÃ", summaryXa.tongHo, 1.0, summaryXa.totalMauChinhThucXa],
        [],
        ["2. BẢNG TỔNG HỢP MẪU VÀ PHÂN LOẠI NÔNG NGHIỆP TỪNG ĐỊA BÀN TRONG XÃ"],
        [
          "STT", 
          "Mã Sheet", 
          "Tên địa bàn", 
          "Tổng số hộ bảng kê", 
          "Số hộ Nông nghiệp (L1)", 
          "Tỷ lệ NN (%)", 
          "Đặc thù địa bàn", 
          "Hộ chính thức", 
          "Hộ dự phòng", 
          "Tổng mẫu", 
          "Loại 1 (CT)", 
          "Loại 2 (CT)", 
          "Loại 3 (CT)", 
          "Loại 4 (CT)"
        ]
      ];

      sheets.forEach((sn, idx) => {
        const db = diaBanMap[sn];
        const hoNN = db.loaiCounts[1] || 0;
        const tyLeNN = db.originalCount > 0 ? (hoNN / db.originalCount) * 100 : 0;
        const dacThu = hoNN === 0 ? "KHÔNG CÓ NÔNG NGHIỆP (0 hộ)" : "Có nông nghiệp";

        const ct1 = db.chinhThuc.filter(h => h.loaiHo === 1).length;
        const ct2 = db.chinhThuc.filter(h => h.loaiHo === 2).length;
        const ct3 = db.chinhThuc.filter(h => h.loaiHo === 3).length;
        const ct4 = db.chinhThuc.filter(h => h.loaiHo === 4).length;
        summaryRows.push([
          idx + 1,
          sn,
          db.tenDiaBan || sn,
          db.originalCount,
          hoNN,
          Number(tyLeNN.toFixed(1)),
          dacThu,
          db.chinhThuc.length,
          db.duPhong.length,
          db.ketQua.length,
          ct1,
          ct2,
          ct3,
          ct4
        ]);
      });

      // Bảng 3: Chi tiết các địa bàn không có nông nghiệp (nếu có)
      const diaBanKhongNNList = sheets.filter(sn => (diaBanMap[sn].loaiCounts[1] || 0) === 0);
      summaryRows.push([]);
      summaryRows.push([`3. DANH SÁCH CÁC ĐỊA BÀN ĐẶC THÙ KHÔNG CÓ HỘ NÔNG NGHIỆP (Tổng cộng: ${diaBanKhongNNList.length} địa bàn)`]);
      summaryRows.push([
        "STT", 
        "Mã Sheet", 
        "Tên địa bàn", 
        "Tổng số hộ", 
        "Hộ Công nghiệp (L2)", 
        "Hộ Dịch vụ (L3)", 
        "Hộ Lương/Khác (L4)", 
        "Nguyên tắc phân bổ mẫu", 
        "Ghi chú kiểm tra"
      ]);

      if (diaBanKhongNNList.length === 0) {
        summaryRows.push(["-", "-", "Toàn bộ địa bàn trong xã đều có hộ nông nghiệp", "-", "-", "-", "-", "-", "Chuẩn"]);
      } else {
        diaBanKhongNNList.forEach((sn, kIdx) => {
          const db = diaBanMap[sn];
          const loaiCoMat = [2, 3, 4].filter((l) => (db.loaiCounts[l] || 0) > 0);
          const loaiNhieuNhat = [...loaiCoMat].sort((a, b) => (db.loaiCounts[b] || 0) - (db.loaiCounts[a] || 0))[0] || 3;
          const tenLoaiNhieu = LOAI_HO_META[loaiNhieuNhat]?.ten || `Loại ${loaiNhieuNhat}`;
          const soHoNhieu = db.loaiCounts[loaiNhieuNhat] || 0;

          const quyTac = loaiCoMat.length === 3
            ? `10 hộ CT phân bổ theo L2, 3, 4; 4 hộ DP gồm 1 hộ L2, 1 hộ L4 và 2 hộ ${tenLoaiNhieu} (gồm 1 đại diện + 1 bù NN do có ${soHoNhieu} hộ)`
            : `10 hộ CT và 4 hộ DP phân bổ theo đúng tỷ lệ các loại hộ hiện có trong địa bàn (${loaiCoMat.map(l => `L${l}`).join(', ')})`;

          summaryRows.push([
            kIdx + 1,
            sn,
            db.tenDiaBan || sn,
            db.originalCount,
            db.loaiCounts[2] || 0,
            db.loaiCounts[3] || 0,
            db.loaiCounts[4] || 0,
            quyTac,
            "Đã kiểm tra không có nông nghiệp (0 hộ)"
          ]);
        });
      }

      const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
      XLSX.utils.book_append_sheet(wb, wsSummary, "00_TongHop_ToanXa");

      // Các sheet địa bàn: Định dạng chuẩn theo phiếu điều tra trong ảnh
      sheets.forEach((sn) => {
        const db = diaBanMap[sn];
        const dbRows: any[][] = [
          ["BẢNG KÊ HỘ CHỌN MẪU ĐIỀU TRA", "", "", "", "", "", "", ""],
          ["ĐIỀU TRA THU NHẬP BÌNH QUÂN ĐẦU NGƯỜI CẤP XÃ", "", "", "", "", "", "", ""],
          [`Tỉnh: ${db.tinh || tinhName}`, "", "", `Xã/phường: ${db.xa || xaName}`, "", "", "", ""],
          [`Tên địa bàn: ${db.tenDiaBan || sn}`, "", "", `(Gồm ${db.chinhThuc.length} hộ chính thức và ${db.duPhong.length} hộ dự phòng)`, "", "", "", ""],
          [],
          [
            "STT Mẫu",
            "Trạng thái mẫu",
            "STT gốc",
            "Họ và tên chủ hộ",
            "Địa chỉ của hộ",
            "Số nhân khẩu",
            "Nguồn thu lớn nhất (1-4)",
            "Phân loại ngành",
            "Ghi chú (SĐT)"
          ],
          // Dòng ký hiệu theo đúng quy chuẩn biểu mẫu
          ["A", "B", "C", "D", "E", "1", "2", "3", "4"]
        ];

        db.ketQua.forEach((h, idx) => {
          let note = h.ghiChu || "";
          if (h.trangThaiMau === "DỰ PHÒNG") {
            const dpText = `Dự phòng thay thế Loại ${h.loaiHo} (${h.tenLoaiHo})`;
            note = note ? `${note} | ${dpText}` : dpText;
          } else if (h.thayTheChoHo) {
            note = note ? `${note} | ${h.thayTheChoHo}` : h.thayTheChoHo;
          }
          dbRows.push([
            idx + 1,
            h.trangThaiMau,
            h.sttGoc,
            h.hoTen,
            h.diaChi,
            h.nhanKhau,
            h.loaiHo,
            h.tenLoaiHo,
            note
          ]);
        });

        const wsDb = XLSX.utils.aoa_to_sheet(dbRows);
        const cleanSheetName = sn.replace(/[:\\/?*\[\]]/g, "_").slice(0, 31);
        XLSX.utils.book_append_sheet(wb, wsDb, cleanSheetName);
      });

      const safeXaName = (communeMeta.xa || "Xa").replace(/\s+/g, "_");
      XLSX.writeFile(wb, `Ket_Qua_Phan_Bo_Ho_${safeXaName}_${sheets.length}DiaBan.xlsx`);
    } catch (err: any) {
      alert(`Lỗi khi xuất file Excel: ${err?.message || "Không xác định"}`);
    }
  };

  // Dữ liệu hiển thị cho địa bàn hiện tại
  const currentDb = diaBanMap[activeSheetTab];
  const filteredHouseholds = useMemo(() => {
    if (!currentDb) return [];
    return currentDb.ketQua.filter((h) => {
      if (filterLoai !== "all" && String(h.loaiHo) !== filterLoai) return false;
      if (filterTrangThai !== "all" && h.trangThaiMau !== filterTrangThai) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const matchName = h.hoTen.toLowerCase().includes(q);
        const matchAddr = h.diaChi.toLowerCase().includes(q);
        const matchStt = String(h.sttGoc).includes(q);
        const matchPhone = h.ghiChu.toLowerCase().includes(q);
        if (!matchName && !matchAddr && !matchStt && !matchPhone) return false;
      }
      return true;
    });
  }, [currentDb, filterLoai, filterTrangThai, searchQuery]);

  return (
    <div className="space-y-4 font-sans text-slate-800 pb-12">
      {/* KHỐI 1: TIÊU ĐỀ VÀ 2 NÚT THAO TÁC */}
      <div className="bg-gradient-to-r from-emerald-900 via-teal-900 to-slate-900 text-white px-5 py-4 rounded-2xl shadow-lg border border-emerald-700/40 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <h1 className="text-base sm:text-lg font-black tracking-tight text-white flex items-center gap-2">
          <Shuffle className="w-5 h-5 text-emerald-300 shrink-0" />
          Chọn mẫu hộ điều tra thu nhập xã{communeMeta.xa ? `: ${communeMeta.xa}` : ""}
        </h1>

        <div className="flex items-center gap-2.5 flex-wrap shrink-0">
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="px-3.5 py-2 rounded-xl text-xs font-bold bg-white/15 hover:bg-white/25 text-white border border-white/30 backdrop-blur-md transition-all flex items-center gap-1.5 cursor-pointer shadow-xs"
            title="Nhấp để chọn file Excel bảng kê"
          >
            <Upload className="w-4 h-4 text-emerald-300" />
            Nạp dữ liệu
          </button>
          
          <button
            type="button"
            onClick={handleExportExcel}
            disabled={Object.keys(diaBanMap).length === 0}
            className="px-4 py-2 rounded-xl text-xs font-black bg-emerald-400 hover:bg-emerald-300 disabled:opacity-40 disabled:cursor-not-allowed text-slate-950 shadow-md transition-all flex items-center gap-1.5 cursor-pointer"
            title="Xuất kết quả phân bổ mẫu ra file Excel"
          >
            <Download className="w-4 h-4" />
            Xuất kết quả
          </button>
        </div>
      </div>

      {/* KHỐI 2: NẠP DỮ LIỆU & ĐỊNH MỨC MẪU */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
        {/* Khối Upload File */}
        <div className="lg:col-span-7 bg-white p-4 rounded-2xl border border-slate-200 shadow-xs space-y-2.5">
          <div className="flex items-center justify-between">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-2">
              <Upload className="w-4 h-4 text-emerald-600" />
              Nạp file bảng kê Excel
            </label>
            {fileName && (
              <span className="text-[11px] font-mono font-semibold text-emerald-700 bg-emerald-50 px-2.5 py-0.5 rounded-md border border-emerald-200 truncate max-w-xs">
                {fileName}
              </span>
            )}
          </div>

          <label className="border-2 border-dashed border-emerald-300 hover:border-emerald-500 bg-emerald-50/40 hover:bg-emerald-50/70 p-4 rounded-xl cursor-pointer flex flex-col items-center justify-center gap-1.5 transition-colors text-center group">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileUpload}
              className="hidden"
            />
            <div className="w-9 h-9 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-700 group-hover:scale-110 transition-transform">
              <FileSpreadsheet className="w-4 h-4" />
            </div>
            <p className="text-xs font-bold text-slate-800">
              Nhấp chuột để chọn file Excel của xã hoặc kéo thả vào đây
            </p>
          </label>

          {errorMessage && (
            <div className="p-2.5 bg-rose-50 border border-rose-200 text-rose-700 rounded-xl text-xs flex items-start gap-2">
              <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5" />
              <span>{errorMessage}</span>
            </div>
          )}
        </div>

        {/* Khối Định Mức Mẫu */}
        <div className="lg:col-span-5 bg-white p-4 rounded-2xl border border-slate-200 shadow-xs flex flex-col justify-between space-y-3">
          <div>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-1.5">
                <Sliders className="w-4 h-4 text-sky-600" />
                Định mức lấy mẫu mỗi địa bàn
              </span>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Số hộ chính thức / ĐB:
                </label>
                <input
                  type="number"
                  min="1"
                  max="50"
                  value={targetChinhThuc}
                  onChange={(e) => setTargetChinhThuc(Math.max(1, parseInt(e.target.value) || 10))}
                  className="w-full text-xs font-bold font-mono px-3 py-1.5 border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500 bg-slate-50"
                />
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Số hộ dự phòng / ĐB:
                </label>
                <input
                  type="number"
                  min="0"
                  max="20"
                  value={targetDuPhong}
                  onChange={(e) => setTargetDuPhong(Math.max(0, parseInt(e.target.value) || 4))}
                  className="w-full text-xs font-bold font-mono px-3 py-1.5 border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500 bg-slate-50"
                />
              </div>
            </div>
          </div>

          <div className="pt-2 border-t border-slate-100 flex items-center gap-2">
            <button
              onClick={handleReshuffle}
              disabled={!rawWorkbook || isProcessing}
              className="w-full py-2 px-3 bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-200 rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 transition-colors cursor-pointer disabled:opacity-50"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isProcessing ? "animate-spin" : ""}`} />
              Đổi đợt ngẫu nhiên (Lần #{randomSeed})
            </button>
          </div>
        </div>
      </div>

      {/* KHỐI 3: TỔNG HỢP CƠ CẤU TOÀN XÃ VÀ 4 LOẠI HỘ */}
      {summaryXa.tongHo > 0 && (
        <div className="space-y-3">
          <div className="flex items-center justify-between flex-wrap gap-2">
            <h2 className="text-xs font-extrabold text-slate-900 uppercase tracking-wide flex items-center gap-2">
              <Database className="w-4 h-4 text-emerald-700" />
              Tổng hợp cơ cấu toàn xã{communeMeta.xa ? `: ${communeMeta.xa}` : ""}
            </h2>
            <div className="text-xs text-slate-600 font-semibold bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200">
              Tổng số: <strong>{summaryXa.tongHo.toLocaleString("vi-VN")} hộ</strong> ({summaryXa.numSheets} địa bàn) | Chính thức: <strong className="text-emerald-700">{summaryXa.totalMauChinhThucXa} hộ</strong> | Dự phòng: <strong className="text-amber-700">{summaryXa.tongDuPhong} hộ</strong>
            </div>
          </div>

          {/* 4 Cards Cơ Cấu 4 Loại Hộ */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
            {[1, 2, 3, 4].map((l) => {
              const meta = LOAI_HO_META[l];
              const count = summaryXa.tongLoai[l] || 0;
              const pct = summaryXa.tongHo ? (count / summaryXa.tongHo) * 100 : 0;
              const quota = summaryXa.mauChiaLoai[l] || 0;

              return (
                <div
                  key={l}
                  className="bg-white p-3 rounded-xl border border-slate-200 shadow-2xs space-y-1.5 hover:border-emerald-300 transition-colors"
                >
                  <div className="flex items-center justify-between">
                    <span className={`text-[10px] font-extrabold px-2 py-0.5 rounded ${meta.badgeColor}`}>
                      Loại {l}
                    </span>
                    <span className="text-xs font-extrabold font-mono text-slate-700">
                      {pct.toFixed(2)}%
                    </span>
                  </div>

                  <h3 className="text-xs font-bold text-slate-900 truncate" title={meta.ten}>
                    {meta.ten}
                  </h3>

                  <div className="pt-1.5 border-t border-slate-100 flex items-center justify-between text-xs">
                    <span className="text-slate-500">Hộ bảng kê:</span>
                    <strong className="font-mono text-slate-800">{count.toLocaleString("vi-VN")} hộ</strong>
                  </div>

                  <div className="flex items-center justify-between text-xs bg-emerald-50/60 p-1.5 rounded-lg border border-emerald-100">
                    <span className="text-emerald-800 font-bold text-[11px]">Mẫu toàn xã:</span>
                    <strong className="font-mono text-emerald-800 font-extrabold">{quota} hộ</strong>
                  </div>

                  <div className="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                    <div
                      className="bg-emerald-600 h-full rounded-full transition-all duration-500"
                      style={{ width: `${Math.min(pct, 100)}%` }}
                    />
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* KHỐI 4: DANH SÁCH HỘ VÀ LỌC THEO ĐỊA BÀN ĐÃ CHỌN MẪU */}
      {Object.keys(diaBanMap).length > 0 && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden space-y-0">
          {/* Header Bảng: Lọc & Tìm Kiếm */}
          <div className="p-3 border-b border-slate-200 bg-slate-50/70 flex flex-col md:flex-row md:items-center justify-between gap-2.5">
            <div className="flex items-center gap-2">
              <Layers className="w-4 h-4 text-emerald-700" />
              <h3 className="text-xs font-bold uppercase tracking-wider text-slate-800">
                Danh sách hộ theo địa bàn
              </h3>
            </div>

            <div className="flex items-center gap-2 flex-wrap">
              <div className="relative">
                <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  placeholder="Tìm chủ hộ, ĐT, địa chỉ..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="pl-8 pr-3 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-500 bg-white w-44"
                />
              </div>

              <select
                value={filterLoai}
                onChange={(e) => setFilterLoai(e.target.value)}
                className="text-xs border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white font-medium focus:outline-none"
              >
                <option value="all">Tất cả loại ngành</option>
                <option value="1">Loại 1: Nông lâm thủy sản</option>
                <option value="2">Loại 2: Công nghiệp xây dựng</option>
                <option value="3">Loại 3: Thương mại dịch vụ</option>
                <option value="4">Loại 4: Nguồn khác / Lương</option>
              </select>

              <select
                value={filterTrangThai}
                onChange={(e) => setFilterTrangThai(e.target.value)}
                className="text-xs border border-slate-300 rounded-lg px-2.5 py-1.5 bg-white font-medium focus:outline-none"
              >
                <option value="all">Tất cả trạng thái</option>
                <option value="CHÍNH THỨC">Chính Thức</option>
                <option value="DỰ PHÒNG">Dự Phòng</option>
                <option value="MẤT MẪU (ĐÃ THAY)">Đã Thay Mẫu</option>
              </select>
            </div>
          </div>

          {/* Thông báo thay thế hộ mất mẫu (nếu có) */}
          {replaceNotice && (
            <div
              className={`px-4 py-2.5 border-b text-xs flex items-center justify-between gap-3 ${
                replaceNotice.type === "success"
                  ? "bg-emerald-50 border-emerald-200 text-emerald-900"
                  : "bg-amber-50 border-amber-200 text-amber-900"
              }`}
            >
              <div className="flex items-center gap-2">
                {replaceNotice.type === "success" ? (
                  <ArrowRightLeft className="w-4 h-4 text-emerald-600 shrink-0" />
                ) : (
                  <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                )}
                <span>{replaceNotice.message}</span>
              </div>
              <button
                type="button"
                onClick={() => setReplaceNotice(null)}
                className="text-slate-400 hover:text-slate-700 font-bold px-1.5 py-0.5 cursor-pointer text-xs"
              >
                ✕
              </button>
            </div>
          )}

          {/* Dải Nút Lọc Theo Địa Bàn Đã Chọn Mẫu */}
          <div className="px-3 py-2 bg-slate-100/70 border-b border-slate-200 flex items-center justify-between gap-3 flex-wrap">
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-slate-600 shrink-0">Lọc địa bàn:</span>
              <select
                value={activeSheetTab}
                onChange={(e) => setActiveSheetTab(e.target.value)}
                className="text-xs font-bold text-slate-800 bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:border-emerald-600 cursor-pointer shadow-2xs"
              >
                {Object.keys(diaBanMap).map((sn) => {
                  const db = diaBanMap[sn];
                  return (
                    <option key={sn} value={sn}>
                      {db.tenDiaBan || sn}
                    </option>
                  );
                })}
              </select>
            </div>

            {/* Dải nút tên địa bàn ngắn gọn */}
            <div className="flex items-center gap-1.5 overflow-x-auto max-w-full py-0.5 custom-scrollbar">
              {Object.keys(diaBanMap).map((sn) => {
                const db = diaBanMap[sn];
                const isActive = activeSheetTab === sn;
                const shortName = db.tenDiaBan || sn;
                return (
                  <button
                    key={sn}
                    type="button"
                    onClick={() => setActiveSheetTab(sn)}
                    className={`px-2.5 py-1 rounded-md text-xs font-bold transition-all shrink-0 cursor-pointer ${
                      isActive
                        ? "bg-emerald-700 text-white shadow-2xs"
                        : "bg-white text-slate-700 hover:bg-slate-200 border border-slate-200"
                    }`}
                  >
                    {shortName}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Thông tin địa bàn đang chọn */}
          {currentDb && (
            <div className="px-4 py-2 bg-emerald-50/40 border-b border-emerald-100 flex items-center justify-between text-xs text-slate-700 flex-wrap gap-2">
              <div>
                Địa bàn: <strong className="text-emerald-950 font-bold">{currentDb.tenDiaBan || currentDb.sheetName}</strong>
                {" | "}Số hộ: <strong>{currentDb.originalCount}</strong>
                {" | "}Đã chọn: <strong className="text-emerald-700">{currentDb.chinhThuc.length} chính thức</strong>, <strong className="text-amber-700">{currentDb.duPhong.length} dự phòng</strong>
              </div>
            </div>
          )}

          {/* Bảng Danh Sách Hộ (Đúng mẫu phiếu bảng kê điều tra) */}
          <div className="overflow-x-auto max-h-96 overflow-y-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200 sticky top-0 z-10">
                <tr>
                  <th className="py-2 px-3 w-14 text-center border-r border-slate-200"># Mẫu</th>
                  <th className="py-2 px-3 w-28 text-center border-r border-slate-200">Trạng Thái</th>
                  <th className="py-2 px-3 w-16 text-center border-r border-slate-200">STT Gốc</th>
                  <th className="py-2 px-3 border-r border-slate-200 min-w-44">Họ và Tên Chủ Hộ</th>
                  <th className="py-2 px-3 border-r border-slate-200 min-w-40">Địa Chỉ Của Hộ</th>
                  <th className="py-2 px-3 w-24 text-center border-r border-slate-200">Số Nhân Khẩu</th>
                  <th className="py-2 px-3 w-40 border-r border-slate-200">Nguồn Thu Lớn Nhất</th>
                  <th className="py-2 px-3 border-r border-slate-200 min-w-32">Ghi Chú / SĐT</th>
                  <th className="py-2 px-3 w-28 text-center">Thao Tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredHouseholds.length === 0 ? (
                  <tr>
                    <td colSpan={9} className="py-8 text-center text-slate-400 italic">
                      Không có hộ nào phù hợp với bộ lọc tìm kiếm.
                    </td>
                  </tr>
                ) : (
                  filteredHouseholds.map((h, idx) => {
                    const isChinhThuc = h.trangThaiMau === "CHÍNH THỨC";
                    const isDuPhong = h.trangThaiMau === "DỰ PHÒNG";
                    const isMatMau = h.trangThaiMau === "MẤT MẪU (ĐÃ THAY)";

                    return (
                      <tr
                        key={idx}
                        className={`hover:bg-slate-50 transition-colors ${
                          isChinhThuc ? "bg-white" : isDuPhong ? "bg-amber-50/25" : "bg-slate-100/60 opacity-70"
                        }`}
                      >
                        <td className="py-2 px-3 text-center font-mono font-bold text-slate-500 border-r border-slate-200">
                          {h.sttMau || idx + 1}
                        </td>
                        <td className="py-2 px-3 text-center border-r border-slate-200">
                          <span
                            className={`inline-block px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-wider ${
                              isChinhThuc
                                ? "bg-emerald-100 text-emerald-800 border border-emerald-300"
                                : isDuPhong
                                ? "bg-amber-100 text-amber-800 border border-amber-300"
                                : "bg-rose-100 text-rose-800 border border-rose-300"
                            }`}
                          >
                            {isMatMau ? "MẤT MẪU" : h.trangThaiMau}
                          </span>
                          {isDuPhong && (
                            <div className="text-[9px] text-amber-700 font-semibold mt-0.5">
                              Thay L{h.loaiHo}
                            </div>
                          )}
                        </td>
                        <td className="py-2 px-3 text-center font-mono text-slate-600 border-r border-slate-200">
                          {h.sttGoc}
                        </td>
                        <td className="py-2 px-3 border-r border-slate-200">
                          <div className={`font-bold ${isMatMau ? "line-through text-slate-500" : "text-slate-900"}`}>
                            {h.hoTen}
                          </div>
                          {h.thayTheChoHo && (
                            <div className="text-[10px] font-medium text-emerald-700 flex items-center gap-1 mt-0.5">
                              <ArrowRightLeft className="w-2.5 h-2.5 shrink-0" />
                              <span>{h.thayTheChoHo}</span>
                            </div>
                          )}
                        </td>
                        <td className="py-2 px-3 text-slate-600 border-r border-slate-200">
                          {h.diaChi || "—"}
                        </td>
                        <td className="py-2 px-3 text-center font-mono text-slate-700 border-r border-slate-200">
                          {h.nhanKhau}
                        </td>
                        <td className="py-2 px-3 border-r border-slate-200">
                          <span
                            className={`inline-block text-[11px] font-semibold px-2 py-0.5 rounded ${
                              LOAI_HO_META[h.loaiHo]?.bgSoft || "bg-slate-100 text-slate-700 border border-slate-200"
                            }`}
                          >
                            Loại {h.loaiHo}: {h.tenLoaiHo}
                          </span>
                        </td>
                        <td className="py-2 px-3 text-slate-700 text-xs font-mono border-r border-slate-200">
                          {h.ghiChu ? (
                            <span className="flex items-center gap-1 text-slate-800">
                              {h.ghiChu.match(/^\d+$/) && <Phone className="w-3 h-3 text-slate-400 shrink-0" />}
                              {h.ghiChu}
                            </span>
                          ) : (
                            <span className="text-slate-400 font-sans italic">—</span>
                          )}
                        </td>
                        <td className="py-2 px-3 text-center">
                          {isChinhThuc && (
                            <button
                              type="button"
                              onClick={() => handleReplaceHousehold(activeSheetTab, h.sttGoc)}
                              className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-bold transition-all cursor-pointer shadow-2xs"
                              title="Khi hộ này vắng nhà hoặc từ chối, bấm để tự động thay bằng hộ dự phòng cùng loại nguồn thu trong địa bàn"
                            >
                              <ArrowRightLeft className="w-3 h-3 text-rose-600" />
                              <span>Đổi mẫu</span>
                            </button>
                          )}
                          {isDuPhong && (
                            <span className="text-[10px] font-medium text-amber-700 bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                              Dự phòng
                            </span>
                          )}
                          {isMatMau && (
                            <span className="text-[10px] text-slate-400 italic">
                              Đã đổi
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};

export default PhanBoHoXa;
