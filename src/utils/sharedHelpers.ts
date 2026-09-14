// Utilities & Helper functions for VTong Data Platform

/**
 * An toàn phân tích số từ chuỗi hoặc giá trị bất kỳ.
 * Hỗ trợ định dạng số Việt Nam (dấu chấm phân cách hàng nghìn, phẩy thập phân),
 * định dạng quốc tế (phẩy phân cách hàng nghìn, chấm thập phân), và số âm trong ngoặc (100).
 */
export function parseRobustNumber(val: any): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;

  let str = String(val).trim();
  if (!str) return 0;

  // Xử lý số âm trong ngoặc (ví dụ: (500) -> -500)
  let isNegative = false;
  if (str.startsWith("(") && str.endsWith(")")) {
    isNegative = true;
    str = str.slice(1, -1).trim();
  } else if (str.startsWith("-")) {
    isNegative = true;
    str = str.slice(1).trim();
  }

  // Loại bỏ các ký tự tiền tệ hoặc đơn vị: đ, vnđ, vnd, usd, %, triệu, tỷ...
  str = str.replace(/[^\d.,]/gi, "").trim();
  if (!str) return 0;

  // Nhận diện kiểu phân cách:
  // Trường hợp 1: Có cả . và ,
  const lastDot = str.lastIndexOf(".");
  const lastComma = str.lastIndexOf(",");

  if (lastDot > -1 && lastComma > -1) {
    if (lastComma > lastDot) {
      // Kiểu VN/EU: 1.234.567,89 -> xóa chấm, đổi phẩy thành chấm
      str = str.replace(/\./g, "").replace(",", ".");
    } else {
      // Kiểu US/UK: 1,234,567.89 -> xóa phẩy
      str = str.replace(/,/g, "");
    }
  } else if (lastComma > -1) {
    // Chỉ có dấu phẩy
    const commaCount = (str.match(/,/g) || []).length;
    if (commaCount > 1) {
      // Nhiều dấu phẩy: 1,234,567 -> xóa phẩy
      str = str.replace(/,/g, "");
    } else {
      // 1 dấu phẩy: nếu phần sau dấu phẩy có đúng 3 chữ số và đứng sau số lớn -> có thể là phân cách ngàn
      const parts = str.split(",");
      if (parts[1] && parts[1].length === 3 && parts[0].length >= 1 && parts[0].length <= 3) {
        str = str.replace(",", "");
      } else {
        str = str.replace(",", ".");
      }
    }
  } else if (lastDot > -1) {
    // Chỉ có dấu chấm
    const dotCount = (str.match(/\./g) || []).length;
    if (dotCount > 1) {
      // Nhiều dấu chấm: 1.234.567 -> kiểu VN ngàn -> xóa chấm
      str = str.replace(/\./g, "");
    }
    // Nếu chỉ có 1 dấu chấm: giữ nguyên chuẩn JS
  }

  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  return isNegative ? -num : num;
}

/**
 * Phân tích nội dung văn bản CSV / TSV thành mảng đối tượng.
 */
export function parseCSV(text: string): Record<string, any>[] {
  if (!text || !text.trim()) return [];

  const lines = text.split(/\r\n|\n|\r/).filter(l => l.trim().length > 0);
  if (lines.length === 0) return [];

  // Tự động phát hiện dấu phân cách (phẩy, chấm phẩy, tab, hoặc |)
  const firstLine = lines[0];
  const delimiters = [",", ";", "\t", "|"];
  let chosenDelim = ",";
  let maxCount = -1;

  for (const d of delimiters) {
    const count = (firstLine.match(new RegExp(`\\${d}`, "g")) || []).length;
    if (count > maxCount) {
      maxCount = count;
      chosenDelim = d;
    }
  }

  // Hàm tách 1 dòng CSV có tính đến dấu ngoặc kép "..."
  const parseLine = (line: string): string[] => {
    const result: string[] = [];
    let current = "";
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i++; // Bỏ qua dấu ngoặc kép thoát
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === chosenDelim && !inQuotes) {
        result.push(current.trim());
        current = "";
      } else {
        current += char;
      }
    }
    result.push(current.trim());
    return result;
  };

  const headers = parseLine(lines[0]).map((h, i) => h || `Cột_${i + 1}`);
  const data: Record<string, any>[] = [];

  for (let i = 1; i < lines.length; i++) {
    const values = parseLine(lines[i]);
    if (values.every(v => !v)) continue; // Bỏ qua dòng trống

    const row: Record<string, any> = {};
    headers.forEach((h, idx) => {
      row[h] = values[idx] !== undefined ? values[idx] : "";
    });
    data.push(row);
  }

  return data;
}

/**
 * Tự động chẩn đoán dòng tiêu đề thông minh từ mảng 2 chiều đọc từ Excel (XLSX).
 * Tìm dòng có nhiều chuỗi có nghĩa nhất làm header, bỏ qua các dòng tiêu đề báo cáo phía trên.
 */
export function parse2DArrayWithSmartHeader(rawRows: any[][]): { data: Record<string, any>[]; columns: string[] } {
  if (!rawRows || rawRows.length === 0) {
    return { data: [], columns: [] };
  }

  // Tìm dòng tiêu đề tốt nhất trong tối đa 15 dòng đầu tiên
  let bestHeaderIndex = 0;
  let maxScore = -1;

  const maxSearchRows = Math.min(rawRows.length, 15);
  for (let r = 0; r < maxSearchRows; r++) {
    const row = rawRows[r];
    if (!Array.isArray(row)) continue;

    let textCellsCount = 0;
    let totalLength = 0;

    row.forEach(cell => {
      if (cell !== null && cell !== undefined) {
        const str = String(cell).trim();
        // Kiểm tra xem cell có phải chuỗi không phải số thuần túy
        if (str.length > 0 && isNaN(Number(str))) {
          textCellsCount++;
          totalLength += str.length;
        }
      }
    });

    // Điểm số ưu tiên số lượng cột tiêu đề dạng chữ
    const score = textCellsCount * 10 + (totalLength > 0 ? 5 : 0);
    if (score > maxScore) {
      maxScore = score;
      bestHeaderIndex = r;
    }
  }

  const rawHeaderRow = rawRows[bestHeaderIndex] || [];
  const colNames: string[] = [];
  const colSeen = new Map<string, number>();

  rawHeaderRow.forEach((cell, idx) => {
    let name = (cell !== null && cell !== undefined ? String(cell).trim() : "") || `Cột_${idx + 1}`;
    name = name.replace(/\r?\n|\r/g, " "); // Xóa xuống dòng trong tên cột

    if (colSeen.has(name)) {
      const count = colSeen.get(name)! + 1;
      colSeen.set(name, count);
      colNames.push(`${name}_${count}`);
    } else {
      colSeen.set(name, 1);
      colNames.push(name);
    }
  });

  const data: Record<string, any>[] = [];
  for (let r = bestHeaderIndex + 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!Array.isArray(row)) continue;

    // Kiểm tra dòng có dữ liệu không
    const hasData = row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== "");
    if (!hasData) continue;

    const rowObj: Record<string, any> = {};
    colNames.forEach((colName, colIdx) => {
      const val = row[colIdx];
      rowObj[colName] = val !== null && val !== undefined ? val : "";
    });
    data.push(rowObj);
  }

  return { data, columns: colNames };
}

/**
 * Làm đẹp tên cột: viết hoa chữ cái đầu, xóa khoảng trắng thừa.
 */
export function beautifyColumnName(col: string): string {
  if (!col) return "";
  let clean = col.replace(/[_\-]+/g, " ").trim();
  return clean
    .split(/\s+/)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(" ");
}

/**
 * Chấm điểm độ tương thích của tên cột cho vai trò cụ thể trong ngành thống kê.
 */
export function scoreColumnForRole(
  col: string,
  role: "mota" | "manganh" | "xa" | "doanhthu" | "laodong" | "idCol"
): number {
  if (!col) return 0;
  const c = col.toLowerCase().trim();

  switch (role) {
    case "idCol":
      if (c === "mst" || c === "tax_code" || c === "mã số thuế") return 100;
      if (c.includes("mst") || c.includes("mã thuế") || c.includes("tax")) return 90;
      if (c === "stt" || c === "id" || c === "mã cơ sở" || c === "ma_cs") return 80;
      if (c.includes("mã") && !c.includes("ngành") && !c.includes("xã")) return 60;
      return 0;

    case "mota":
      if (c.includes("mô tả") || c.includes("hoạt động") || c.includes("ngành nghề kinh doanh")) return 100;
      if (c.includes("nội dung") || c.includes("ten_nganh") || c.includes("tên ngành")) return 90;
      if (c.includes("nghề") || c.includes("sản phẩm")) return 70;
      return 0;

    case "manganh":
      if (c === "mã ngành" || c === "manganh" || c === "ma_nganh" || c === "vsic") return 100;
      if (c.includes("mã ngành") || c.includes("manganh") || c.includes("vsic")) return 90;
      if (c.includes("ngành cấp 5") || c.includes("cấp 5") || c.includes("ma_c5")) return 85;
      if (c.includes("ngành") && (c.includes("mã") || c.includes("code"))) return 80;
      return 0;

    case "xa":
      if (c === "mã xã" || c === "maxa" || c === "ma_xa" || c === "xã") return 100;
      if (c.includes("xã") || c.includes("phường") || c.includes("thị trấn")) return 90;
      if (c.includes("địa bàn") || c.includes("huyện") || c.includes("tỉnh")) return 70;
      return 0;

    case "doanhthu":
      if (c === "doanh thu" || c === "doanhthu" || c === "doanh_thu") return 100;
      if (c.includes("doanh thu") || c.includes("doanh số") || c.includes("revenue")) return 90;
      if (c.includes("thu nhập") || c.includes("tiền bán") || c.includes("tổng thu")) return 80;
      return 0;

    case "laodong":
      if (c === "lao động" || c === "laodong" || c === "lao_dong") return 100;
      if (c.includes("lao động") || c.includes("nhân viên") || c.includes("số người")) return 90;
      if (c.includes("lao dong") || c.includes("worker") || c.includes("staff")) return 80;
      return 0;
  }
}

/**
 * Tự động gán duy nhất các vai trò tốt nhất cho danh sách cột.
 */
export function getUniqueRoleAssignments(columns: string[]): Record<string, string> {
  const roles: ("idCol" | "mota" | "manganh" | "xa" | "doanhthu" | "laodong")[] = [
    "idCol",
    "mota",
    "manganh",
    "xa",
    "doanhthu",
    "laodong"
  ];

  const assignments: Record<string, string> = {};
  const assignedRoles = new Set<string>();

  for (const role of roles) {
    let bestCol = "";
    let highestScore = 40; // Ngưỡng tối thiểu

    for (const col of columns) {
      if (assignments[col]) continue; // Cột đã được gán vai trò khác
      const score = scoreColumnForRole(col, role);
      if (score > highestScore) {
        highestScore = score;
        bestCol = col;
      }
    }

    if (bestCol) {
      assignments[bestCol] = role;
      assignedRoles.add(role);
    }
  }

  return assignments;
}

/**
 * Lấy giá trị linh hoạt từ một hàng dựa trên tên cột hoặc tên gần đúng.
 */
export function getFlexibleValue(row: any, colName: string): any {
  if (!row || typeof row !== "object" || !colName) return undefined;

  // Khớp chính xác
  if (row[colName] !== undefined) return row[colName];

  // Khớp không phân biệt hoa thường và khoảng trắng
  const target = colName.toLowerCase().trim();
  for (const key of Object.keys(row)) {
    if (key.toLowerCase().trim() === target) {
      return row[key];
    }
  }

  return undefined;
}

/**
 * Chuẩn hóa biểu thức logic do AI sinh ra để có thể thực thi an toàn trong JavaScript.
 */
export function normalizeAiExpression(expr: string): string {
  if (!expr) return "false";

  let clean = expr.trim();
  // Loại bỏ các tiền tố hoặc giải thích thừa
  if (clean.startsWith("```") && clean.endsWith("```")) {
    clean = clean.replace(/```(?:javascript|js)?/g, "").replace(/```/g, "").trim();
  }

  // Chuyển đổi cú pháp row["Cột"] hoặc row.Cột thành getFlexibleValue(row, "Cột")
  clean = clean.replace(/row\[(["'])(.*?)\1\]/g, 'getFlexibleValue(row, "$2")');
  clean = clean.replace(/row\.([a-zA-Z0-9_\u00C0-\u024F\u1EA0-\u1EF9]+)/g, 'getFlexibleValue(row, "$1")');

  return clean;
}

/**
 * Chấm điểm độ phù hợp của cột để làm khóa chung khi ghép sheet.
 * Ưu tiên cao nhất: Mã Hộ, Mã Cơ Sở, Mã Số Thuế, Mã Định Danh, ID, STT...
 */
export function scoreCommonKeyCandidate(colName: string, coverageRatio: number): number {
  let score = coverageRatio * 100;
  const clean = colName.toLowerCase().replace(/[\s_\-\.]/g, "");

  // Siêu ưu tiên các cột mã định danh đối tượng thống kê & hộ gia đình
  if (/^(maho|idho|sttho|mahogiadinh|sothutuho|sohogiadinh|mahokh|idhouse|householdid)/i.test(clean)) {
    score += 500;
  } else if (/^(macoso|mst|masothue|madoituong|iddoituong|makh|madv)/i.test(clean)) {
    score += 400;
  } else if (/^(id|uuid|key|code|stt|sothutu)/i.test(clean)) {
    score += 300;
  } else if (/(ma|id|code|stt|dinhdanh)/i.test(clean)) {
    score += 200;
  } else if (/(madiaban|maxa|mahuyen|matinh|madb)/i.test(clean)) {
    score += 150;
  }

  // Giảm điểm nặng nếu là các cột thông tin phụ hoặc mô tả dài
  if (/(ghichu|mota|ten|hoten|diachi|ngaysinh|namsinh|gioitinh|quanhe|dienthoai|sdt|email|stt_tv)/i.test(clean)) {
    score -= 100;
  }
  if (clean.startsWith("__empty")) {
    score -= 1000;
  }

  return score;
}

/**
 * Phân tích và tìm các cột chung giữa các sheet được chọn.
 */
export function analyzeWorkbookCommonColumns(
  selectedSheets: string[],
  sheetHeadersMap: Record<string, string[]>
): {
  exactCommonCols: string[];
  partialCols: { name: string; count: number }[];
  allSharedCols: string[];
  bestCommonCol: string;
} {
  if (!selectedSheets || selectedSheets.length === 0) {
    return { exactCommonCols: [], partialCols: [], allSharedCols: [], bestCommonCol: "" };
  }

  // Nếu chỉ có 1 sheet
  if (selectedSheets.length === 1) {
    const cols = (sheetHeadersMap[selectedSheets[0]] || []).filter(c => c && !c.startsWith("__EMPTY"));
    let best = cols[0] || "";
    let highest = -999;
    cols.forEach(c => {
      const s = scoreCommonKeyCandidate(c, 1);
      if (s > highest) {
        highest = s;
        best = c;
      }
    });
    return {
      exactCommonCols: cols,
      partialCols: [],
      allSharedCols: cols,
      bestCommonCol: best
    };
  }

  // Thu thập tần suất xuất hiện của từng tên cột qua các sheet đã chọn
  const colMap = new Map<string, { originalName: string; sheets: Set<string> }>();

  selectedSheets.forEach(sheetName => {
    const cols = sheetHeadersMap[sheetName] || [];
    cols.forEach(col => {
      if (!col || col.startsWith("__EMPTY")) return;
      const key = col.trim().toLowerCase();
      if (!colMap.has(key)) {
        colMap.set(key, { originalName: col.trim(), sheets: new Set() });
      }
      colMap.get(key)!.sheets.add(sheetName);
    });
  });

  const totalSheets = selectedSheets.length;
  const exactCommonCols: string[] = [];
  const partialCols: { name: string; count: number }[] = [];
  const scoredCandidates: { name: string; score: number }[] = [];

  colMap.forEach((val) => {
    const count = val.sheets.size;
    const ratio = count / totalSheets;
    const score = scoreCommonKeyCandidate(val.originalName, ratio);

    if (count === totalSheets) {
      exactCommonCols.push(val.originalName);
      scoredCandidates.push({ name: val.originalName, score });
    } else if (count >= 2) {
      partialCols.push({ name: val.originalName, count });
      scoredCandidates.push({ name: val.originalName, score });
    }
  });

  // Sắp xếp các cột chung tuyệt đối theo điểm ưu tiên
  exactCommonCols.sort((a, b) => {
    const sA = scoreCommonKeyCandidate(a, 1);
    const sB = scoreCommonKeyCandidate(b, 1);
    return sB - sA;
  });

  partialCols.sort((a, b) => b.count - a.count);
  scoredCandidates.sort((a, b) => b.score - a.score);

  const bestCommonCol = scoredCandidates.length > 0 ? scoredCandidates[0].name : (exactCommonCols[0] || "");
  const allSharedCols = Array.from(new Set([...exactCommonCols, ...partialCols.map(p => p.name)]));

  return {
    exactCommonCols,
    partialCols,
    allSharedCols,
    bestCommonCol
  };
}
