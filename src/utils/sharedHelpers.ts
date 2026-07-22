// Shared utilities and helpers to prevent circular dependencies in App.tsx and components

// Hàm hỗ trợ so khớp đàn hồi mềm dẻo và lấy dữ liệu cột từ hàng (row) để chống sai phông chữ bừa bãi hoặc hoa thường lệch lạc từ AI
export function getFlexibleValue(row: any, keyName: string): any {
  if (!row || !keyName) return "";
  if (keyName in row) return row[keyName];

  const cleanKey = keyName.toLowerCase().replace(/\s+/g, "").trim();
  const actualKeys = Object.keys(row);

  // Thử khớp viết thường không khoảng trắng
  for (const k of actualKeys) {
    const cleanK = k.toLowerCase().replace(/\s+/g, "").trim();
    if (cleanK === cleanKey) {
      return row[k];
    }
  }

  // Thử khớp không dấu tiếng Việt
  const stripDiacritics = (str: string) => 
    str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/\s+/g, "");

  const searchNormalized = stripDiacritics(keyName);
  for (const k of actualKeys) {
    if (stripDiacritics(k) === searchNormalized) {
      return row[k];
    }
  }

  return row[keyName] || "";
}

// Hàm chuẩn hóa biểu thức AI sang biểu diễn an toàn qua getFlexibleValue
export function normalizeAiExpression(expr: string): string {
  if (!expr) return "";
  let clean = expr;
  
  // 1. Chuyển đổi các cặp row['Cột'] hoặc row["Cột"] thành getFlexibleValue(row, 'Cột')
  clean = clean.replace(/(?:row|Row)\s*\[\s*['"]([^'"]+)['"]\s*\]/g, "getFlexibleValue(row, '$1')");
  
  // 2. Chuyển đổi các dạng row.TenCol thành getFlexibleValue(row, 'TenCol')
  // Chỉ khớp thuộc tính alpha-numeric bắt đầu bằng chữ cái
  clean = clean.replace(/(?:row|Row)\.([a-zA-Z_][a-zA-Z0-9_]*)/g, "getFlexibleValue(row, '$1')");
  
  return clean;
}

// Bộ phân giải CSV tối ưu hóa cao cho các tệp lớn (như 50MB+), loại bỏ lỗi "Too many properties to enumerate"
export const parseCSV = (rawText: string): any[] => {
  let text = rawText;
  if (text.charCodeAt(0) === 0xFEFF) {
    text = text.substring(1);
  }

  // Tự động phát hiện dấu phân tách (comma, semicolon, tab) từ dòng dữ liệu đầu tiên
  const firstLineEnd = text.indexOf('\n');
  const firstLine = firstLineEnd === -1 ? text : text.substring(0, firstLineEnd);
  
  let delimiter = ',';
  const commaCount = (firstLine.match(/,/g) || []).length;
  const semicolonCount = (firstLine.match(/;/g) || []).length;
  const tabCount = (firstLine.match(/\t/g) || []).length;
  
  if (semicolonCount > commaCount && semicolonCount > tabCount) {
    delimiter = ';';
  } else if (tabCount > commaCount && tabCount > semicolonCount) {
    delimiter = '\t';
  }

  const length = text.length;
  const rows: string[][] = [];
  let currentRow: string[] = [];
  
  let i = 0;
  let inQuotes = false;
  let start = 0;
  
  while (i < length) {
    const char = text[i];
    
    if (char === '"') {
      if (!inQuotes) {
        inQuotes = true;
        start = i + 1; // nhảy qua dấu ngoặc kép mở đầu
      } else {
        // Kiểm tra dấu ngoặc kép trốn (escaped quote "")
        if (i + 1 < length && text[i + 1] === '"') {
          i++; // bỏ qua dấu ngoặc kép thứ hai
        } else {
          inQuotes = false;
        }
      }
    } else if (!inQuotes) {
      if (char === delimiter) {
        let cell = text.substring(start, i);
        // Loại bỏ ngoặc kép bao quanh khi đọc chuỗi trường
        if (text[i - 1] === '"' && text[start - 1] === '"') {
          cell = cell.substring(0, cell.length - 1);
        }
        if (cell.includes('""')) {
          cell = cell.replace(/""/g, '"');
        }
        currentRow.push(cell.trim());
        start = i + 1;
      } else if (char === '\n' || char === '\r') {
        let cell = text.substring(start, i);
        if (text[i - 1] === '"' && text[start - 1] === '"') {
          cell = cell.substring(0, cell.length - 1);
        }
        if (cell.includes('""')) {
          cell = cell.replace(/""/g, '"');
        }
        currentRow.push(cell.trim());
        
        if (currentRow.length > 0 && (currentRow.length > 1 || currentRow[0] !== "")) {
          rows.push(currentRow);
        }
        currentRow = [];
        
        if (char === '\r' && i + 1 < length && text[i + 1] === '\n') {
          i++;
        }
        start = i + 1;
      }
    }
    i++;
  }
  
  if (start < length) {
    let cell = text.substring(start, length);
    if (text[length - 1] === '"' && text[start - 1] === '"') {
      cell = cell.substring(0, cell.length - 1);
    }
    if (cell.includes('""')) {
      cell = cell.replace(/""/g, '"');
    }
    currentRow.push(cell.trim());
  }
  if (currentRow.length > 0 && (currentRow.length > 1 || currentRow[0] !== "")) {
    rows.push(currentRow);
  }

  if (rows.length === 0) return [];

  const headers = rows[0];
  const data: any[] = [];
  for (let idx = 1; idx < rows.length; idx++) {
    const r = rows[idx];
    const obj: any = {};
    let hasData = false;
    for (let c = 0; c < headers.length; c++) {
      const headerName = headers[c] || `Cột ${c + 1}`;
      const val = r[c] !== undefined ? r[c] : "";
      obj[headerName] = val;
      if (val !== "") hasData = true;
    }
    if (hasData) {
      data.push(obj);
    }
  }

  return data;
};

// Hàm tự động chuyển đổi tên cột (tiếng Anh viết tắt, snake_case, camelCase...) thành tiếng Việt thân thiện, rõ nghĩa
export function beautifyColumnName(colName: string): string {
  if (!colName) return "";

  // 1. Tách CamelCase thành khoảng trắng (vd: MaNganhDTV -> Ma Nganh DTV, MaNganhC5 -> Ma Nganh C5)
  let name = colName.replace(/([a-z0-9])([A-Z])/g, "$1 $2");
  
  // 2. Thay thế các ký tự gạch dưới, gạch ngang thành khoảng trắng
  name = name.replace(/[_-]+/g, " ");

  // 3. Chuẩn hóa khoảng trắng dư thừa
  name = name.trim().replace(/\s+/g, " ");

  // 4. Định nghĩa một số cụm từ chính xác trước để bảo vệ ý nghĩa chuyên môn
  const lowerName = name.toLowerCase();
  
  // Áp dụng các từ viết tắt thông dụng trong lĩnh vực kinh tế / thống kê / thuế
  const words = name.split(" ");
  const processedWords = words.map((w, idx) => {
    const wl = w.toLowerCase();
    
    // Áp dụng bộ từ điển việt hóa thông minh cho từng từ hoặc cụm từ
    if (wl === "doanhthu" || wl === "revenue") return "Doanh thu";
    if (wl === "thuan" || wl === "net") return "Thuần";
    if (wl === "doanhso" || wl === "turnover" || wl === "sales") return "Doanh số";
    if (wl === "mota" || wl === "desc" || wl === "description") return "Mô tả";
    if (wl === "manganh" || wl === "vsic") return "Mã ngành";
    if (wl === "laodong" || wl === "labor" || wl === "ld") return "Lao động";
    if (wl === "nhansu" || wl === "staff") return "Nhân sự";
    if (wl === "mast" || wl === "mst" || wl === "tax" || wl === "taxcode") return "Mã số thuế";
    if (wl === "diaban") return "Địa bàn";
    if (wl === "xa" || wl === "commune") return "Xã";
    if (wl === "phuong" || wl === "ward") return "Phường";
    if (wl === "tytrong" || wl === "ratio" || wl === "share" || wl === "ty_trong") return "Tỷ trọng";
    if (wl === "sudung" || wl === "usage" || wl === "su_dung") return "Sử dụng";
    if (wl === "tmdt" || wl === "ecommerce" || wl === "tm_dt") return "TMĐT";
    if (wl === "giaohang" || wl === "delivery" || wl === "giao_hang") return "Giao hàng";
    if (wl === "tindung" || wl === "credit" || wl === "tin_dung") return "Tín dụng";
    if (wl === "bhccdv" || wl === "services") return "BHCCDV";
    if (wl === "hinhthuc" || wl === "type" || wl === "form" || wl === "hinh_thuc") return "Hình thức";
    if (wl === "nganh" || wl === "sector") return "Ngành";
    if (wl === "nghe") return "Nghề";
    if (wl === "dtv") return "DTV";
    if (wl === "stt") return "STT";
    if (wl === "dk") return "ĐK";
    if (wl === "ma" || wl === "code") return "Mã";
    if (wl === "so" || wl === "number") return "Số";
    if (wl === "thue") return "Thuế";
    if (wl === "huyen") return "Huyện";
    if (wl === "tinh") return "Tỉnh";
    
    // Nếu là từ bình thường, viết hoa chữ cái đầu
    if (w.length > 0) {
      // Giữ nguyên các từ viết hoa toàn bộ có từ 2 ký tự trở lên (TMĐT, MST, VSIC, BHCCDV...)
      if (w === w.toUpperCase() && w.length > 1) return w;
      return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    }
    return w;
  });

  let result = processedWords.join(" ");

  // Thêm một số tinh chỉnh tiếng Việt để đọc thuận miệng hơn
  result = result.replace(/Ma Nganh/gi, "Mã ngành");
  result = result.replace(/Ma So Thue/gi, "Mã số thuế");
  result = result.replace(/Doanh Thu/gi, "Doanh thu");
  result = result.replace(/Doanh So/gi, "Doanh số");
  result = result.replace(/Lao Dong/gi, "Lao động");
  result = result.replace(/Nhan Su/gi, "Nhân sự");
  result = result.replace(/Dia Ban/gi, "Địa bàn");
  result = result.replace(/Mo Ta/gi, "Mô tả");
  result = result.replace(/Hinh Thuc/gi, "Hình thức");
  result = result.replace(/Ty Trong/gi, "Tỷ trọng");
  result = result.replace(/Tin Dung/gi, "Tín dụng");
  result = result.replace(/Giao Hang/gi, "Giao hàng");
  result = result.replace(/Su Dung/gi, "Sử dụng");
  result = result.replace(/Nganh Nghe/gi, "Ngành nghề");
  result = result.replace(/Noi Dung/gi, "Nội dung");
  result = result.replace(/Dien Giai/gi, "Diễn giải");
  
  return result;
}

// Tính điểm độ tương khớp của tên cột đối với vai trò hệ thống (để gán nhãn chính xác, tránh gán nhầm lẫn)
export function scoreColumnForRole(colName: string, role: string): number {
  if (!colName || !role) return 0;
  
  const cl = colName.toLowerCase().replace(/[^a-z0-9_]/g, "");
  
  if (role === "idCol") {
    if (/mst|mast|tax|taxcode|masothue|ma_so_thue/i.test(colName)) return 100;
    if (/^id$|_id$|^id_|madn|ma_dn|madoanhnghiep|ma_doanh_nghiep/i.test(colName)) return 50;
    return 0;
  }
  
  if (role === "mota") {
    if (/mota|mo_ta|diengiai|dien_giai|noidung|noi_dung/i.test(colName)) return 100;
    if (/hoatdong|hoat_dong|nganhnghe|nganh_nghe|description|desc|act|activity/i.test(colName)) return 80;
    return 0;
  }
  
  if (role === "manganh") {
    // Không nhận diện nếu là mô tả ngành nghề để tránh nhầm lẫn
    if (/mota|mo_ta|nganhnghe|nganh_nghe/i.test(colName)) return 0;
    if (/manganh|ma_nganh|vsic|ma_vsic|manganhdtv|ma_nganh_dtv/i.test(colName)) return 100;
    if (/industry|sector|ngành|nganh|code/i.test(colName)) return 50;
    return 0;
  }
  
  if (role === "xa") {
    if (/dia_ban_xa|diabanxa|xa_phuong|xaphuong/i.test(colName)) return 100;
    if (/^xa$|_xa$|^phuong$|_phuong$|diaban|dia_ban/i.test(colName)) return 80;
    if (/town|district|ward|commune|khuvuc|khu_vuc/i.test(colName)) return 50;
    return 0;
  }
  
  if (role === "doanhthu") {
    // Các từ khóa tiêu cực cực kỳ nhạy cảm để loại bỏ ngay lập tức (tránh gán nhầm sang hình thức hoặc tỷ trọng)
    if (/tytrong|ty_trong|ratio|share|hinhthuc|hinh_thuc|type|form/i.test(colName)) return -100;
    
    if (/doanhthu|doanh_thu|doanhso|doanh_so|tongthu|tong_thu/i.test(colName)) return 100;
    if (/revenue|turnover|sales|sale|thunhap|thu_nhap|tien|money/i.test(colName)) return 80;
    if (/trigia|tri_gia|giatri|gia_tri/i.test(colName)) return 50;
    return 0;
  }
  
  if (role === "laodong") {
    if (/laodong|lao_dong|nhansu|nhan_su|solaodong|so_lao_dong/i.test(colName)) return 100;
    if (/ld|labor|employment|nhanvien|nhan_vien|songuoi|so_nguoi|staff/i.test(colName)) return 80;
    return 0;
  }
  
  return 0;
}

// Thực hiện ghép nối vai trò thông minh đảm bảo tính DUY NHẤT (Mỗi vai trò chỉ gán cho tối đa 1 cột có điểm cao nhất)
export function getUniqueRoleAssignments(columns: string[]): { [colName: string]: string } {
  const roles = ["idCol", "mota", "manganh", "xa", "doanhthu", "laodong"];
  const assignments: { [colName: string]: string } = {};
  
  const assignedRoles = new Set<string>();
  const assignedCols = new Set<string>();
  
  const candidates: { col: string; role: string; score: number }[] = [];
  
  columns.forEach(col => {
    roles.forEach(role => {
      const score = scoreColumnForRole(col, role);
      if (score > 0) {
        candidates.push({ col, role, score });
      }
    });
  });
  
  // Sắp xếp giảm dần theo điểm số để ưu tiên gán các cặp khớp mạnh nhất
  candidates.sort((a, b) => b.score - a.score);
  
  candidates.forEach(cand => {
    if (!assignedRoles.has(cand.role) && !assignedCols.has(cand.col)) {
      assignments[cand.col] = cand.role;
      assignedRoles.add(cand.role);
      assignedCols.add(cand.col);
    }
  });
  
  return assignments;
}

// Hàm thông minh tự động tìm kiếm dòng chứa tiêu đề (header) của bảng Excel tốt nhất
// Giúp bỏ qua các dòng tiêu đề hành chính, tên báo cáo hoặc các hàng trống ở đầu tệp
export function findBestHeaderRowIndex(rows: any[][]): number {
  if (rows.length === 0) return 0;
  
  // Danh sách từ khóa chuẩn sau dọn dẹp bằng regex
  const cleanStandardKeywords = [
    "mst", "masothue", "madinhdanh", "macoso", "matinh", "matkcs", "maxa", "madiaban", "iddb", "idcoso", "masocs", "madtv", "madoituong",
    "tencoso", "tendoanhnghiep", "tencoso", "doanhnghiep", "tendonvi", "tendoituong",
    "daidien", "chucoso", "nguoidaidien",
    "diachi", "diaban", "xa", "huyen", "tinh",
    "doanhthu", "doanhthuthuan", "trigia", "doanhso",
    "laodong", "songuoi", "nhansu",
    "manganh", "vsic", "manganhchinh", "manghe"
  ];

  let bestIndex = 0;
  let maxScore = -999999;

  // Quét tối đa 20 dòng đầu tiên của file Excel để tìm dòng tiêu đề phù hợp nhất
  const scanLimit = Math.min(rows.length, 20);
  for (let r = 0; r < scanLimit; r++) {
    const row = rows[r];
    if (!row || !Array.isArray(row)) continue;

    let keywordMatches = 0;
    let nonSpaceCells = 0;
    let numericCells = 0;

    row.forEach(cell => {
      if (cell === undefined || cell === null) return;
      const cellStrRaw = String(cell).trim();
      if (!cellStrRaw) return;

      nonSpaceCells++;

      // Kiểm tra xem có phải là ô số thô thuần túy không (không phải mã kết hợp chữ như K133T01)
      // Ô là số thô thuần túy: kiểu number hoặc là chuỗi chỉ gồm chữ số (và không rỗng)
      if (typeof cell === "number") {
        numericCells++;
      } else if (/^\d+$/.test(cellStrRaw)) {
        numericCells++;
      }

      // Dọn dẹp ô để kiểm tra từ khóa
      const cleanedCell = cellStrRaw
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/đ/g, "d")
        .replace(/[^a-z0-9]/g, "");

      if (!cleanedCell) return;

      // Kiểm tra xem ô dọn dẹp có chứa bất kỳ từ khóa chuẩn dọn dẹp nào không
      const matchesKeyword = cleanStandardKeywords.some(keyword => {
        return cleanedCell.includes(keyword) || keyword.includes(cleanedCell);
      });

      if (matchesKeyword) {
        keywordMatches++;
      }
    });

    if (nonSpaceCells === 0) continue;

    // Tính tỷ lệ ô là số thô
    const numericRatio = numericCells / nonSpaceCells;

    // Điểm cơ bản dựa trên số lượng từ khóa cột khớp
    let score = keywordMatches * 20 + nonSpaceCells;

    // PHẠT CỰC NẶNG NẾU LÀ DÒNG DỮ LIỆU CHỨA NHIỀU SỐ THÔ
    // Hàng tiêu đề thực tế không bao giờ chứa nhiều số thô (chỉ có thể là chữ hoặc chữ kết hợp số như A1_1)
    if (numericRatio > 0.05 || numericCells > 2) {
      score -= (numericCells * 50) + 500; // Trừ điểm rất nặng để loại bỏ dòng dữ liệu
    }

    if (score > maxScore) {
      maxScore = score;
      bestIndex = r;
    }
  }

  // Nếu điểm tối đa quá thấp (ví dụ < 5) hoặc không có dòng nào thực sự vượt trội,
  // chọn dòng không trống đầu tiên có chứa ít nhất 2 ô và tỉ lệ số thô cực thấp (< 5%)
  if (maxScore < 5) {
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      const nonSpace = row.filter(cell => String(cell || "").trim() !== "");
      if (nonSpace.length >= 2) {
        // Đếm số lượng số thô trong dòng này
        let numCells = 0;
        nonSpace.forEach(cell => {
          const cellStr = String(cell).trim();
          if (typeof cell === "number" || /^\d+$/.test(cellStr)) {
            numCells++;
          }
        });
        // Chỉ chọn dòng này làm tiêu đề nếu tỉ lệ số thô rất thấp
        if (numCells / nonSpace.length < 0.1) {
          return r;
        }
      }
    }

    // Fallback cuối cùng nếu toàn bộ đều chứa số hoặc không tìm ra dòng nào đáp ứng
    for (let r = 0; r < rows.length; r++) {
      if (rows[r] && rows[r].filter(cell => String(cell || "").trim() !== "").length >= 2) {
        return r;
      }
    }
    return 0;
  }

  return bestIndex;
}

// Hàm phân tích mảng 2D thô thành danh sách các đối tượng JSON với dòng tiêu đề tốt nhất tự động phát hiện
export function parse2DArrayWithSmartHeader(rows: any[][]): { data: any[], columns: string[], headerIndex: number } {
  if (rows.length === 0) {
    return { data: [], columns: [], headerIndex: 0 };
  }

  const bestHeaderIdx = findBestHeaderRowIndex(rows);
  const rawHeaders = rows[bestHeaderIdx] || [];
  
  // Chuẩn hóa tiêu đề và lọc bỏ các cột không có tên tiêu đề hoặc tự đặt tên cột nếu rỗng nhưng có dữ liệu bên dưới
  const columns: string[] = [];
  rawHeaders.forEach((h, cIdx) => {
    const cleanHeader = String(h || "").trim();
    if (cleanHeader) {
      columns.push(cleanHeader);
    } else {
      // Kiểm tra xem có cột rỗng nào mà bên dưới có dữ liệu không
      let hasDataBelow = false;
      const checkLimit = Math.min(rows.length, bestHeaderIdx + 100); // kiểm tra tối đa 100 dòng bên dưới
      for (let r = bestHeaderIdx + 1; r < checkLimit; r++) {
        if (rows[r] && rows[r][cIdx] !== undefined && String(rows[r][cIdx]).trim() !== "") {
          hasDataBelow = true;
          break;
        }
      }
      if (hasDataBelow) {
        columns.push(`Cột ${cIdx + 1}`);
      } else {
        columns.push(""); // Cột trống thực sự
      }
    }
  });

  const data: any[] = [];
  for (let r = bestHeaderIdx + 1; r < rows.length; r++) {
    const rowData = rows[r];
    if (!rowData || !Array.isArray(rowData)) continue;

    // Bỏ qua các dòng trống hoàn toàn hoặc các dòng chỉ chứa các ký tự trắng
    const isRowEmpty = rowData.every(cell => String(cell || "").trim() === "");
    if (isRowEmpty) continue;

    const rowObj: any = {};
    let hasValidData = false;

    columns.forEach((header, cIndex) => {
      if (!header) return; // Bỏ qua cột rỗng thực sự
      const value = rowData[cIndex] !== undefined ? rowData[cIndex] : "";
      rowObj[header] = value;
      if (String(value).trim() !== "") {
        hasValidData = true;
      }
    });

    if (hasValidData) {
      data.push(rowObj);
    }
  }

  // Lọc bỏ tiêu đề rỗng khỏi danh sách columns trả về cho UI
  const validColumns = columns.filter(c => c !== "");

  return {
    data,
    columns: validColumns,
    headerIndex: bestHeaderIdx
  };
}

// Chuyển từ components/sectorRevenueChart.tsx sang đây để App.tsx (và các module khác) có thể
// dùng hàm này ngay lập tức (đồng bộ), không phụ thuộc việc lazy-load component biểu đồ.
export function parseRobustNumber(val: any): number {
  if (val === undefined || val === null) return 0;
  if (typeof val === "number") return val;
  let str = String(val).trim();
  if (!str) return 0;

  str = str.replace(/\s+/g, "").replace(/[đđVNDvnd]/g, "");

  const hasComma = str.includes(",");
  const hasDot = str.includes(".");

  if (hasComma && hasDot) {
    if (str.indexOf(".") < str.indexOf(",")) {
      str = str.replace(/\./g, "").replace(/,/g, ".");
    } else {
      str = str.replace(/,/g, "");
    }
  } else if (hasComma) {
    const commas = (str.match(/,/g) || []).length;
    if (commas > 1) {
      str = str.replace(/,/g, "");
    } else {
      const parts = str.split(",");
      if (parts[1].length === 3) {
        str = str.replace(/,/g, "");
      } else {
        str = str.replace(/,/g, ".");
      }
    }
  } else if (hasDot) {
    const dots = (str.match(/\./g) || []).length;
    if (dots > 1) {
      str = str.replace(/\./g, "");
    } else {
      const parts = str.split(".");
      if (parts[1].length === 3) {
        str = str.replace(/\./g, "");
      }
    }
  }

  const clean = str.replace(/[^0-9.\-]/g, "");
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}

