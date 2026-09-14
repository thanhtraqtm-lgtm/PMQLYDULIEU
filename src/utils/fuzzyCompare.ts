// Helper module for smart string matching, fuzzy comparison, and flexible VLOOKUP

/**
 * Loại bỏ dấu tiếng Việt chuẩn Unicode NFD.
 */
export function removeVietnameseTones(str: string): string {
  if (!str) return "";
  return str
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D");
}

export interface CleanOptions {
  ignoreCase?: boolean;
  ignoreDiacritics?: boolean;
  ignoreSpaces?: boolean;
  ignorePunctuation?: boolean;
}

/**
 * Làm sạch chuỗi theo các tùy chọn linh hoạt:
 * - Bỏ dấu tiếng Việt
 * - Chuyển chữ thường
 * - Bỏ toàn bộ khoảng trắng (kể cả giữa chuỗi)
 * - Bỏ dấu phân cách và ký tự đặc biệt (_ - . / , : ; () [] {})
 */
export function cleanForComparison(val: any, options: CleanOptions = {}): string {
  if (val === null || val === undefined) return "";
  let s = String(val).trim();

  if (options.ignoreDiacritics !== false) {
    s = removeVietnameseTones(s);
  }

  if (options.ignoreCase !== false) {
    s = s.toLowerCase();
  }

  if (options.ignorePunctuation !== false) {
    // Bỏ tất cả dấu câu, gạch dưới, gạch ngang, chấm, phẩy, chéo, ngoặc...
    s = s.replace(/[\s_\-.,;:/\\[\](){}'"`~*+?=!@#$%^&|<>]/g, "");
  } else if (options.ignoreSpaces !== false) {
    // Nếu không bỏ punctuation nhưng bỏ spaces thì bỏ toàn bộ khoảng trắng
    s = s.replace(/\s+/g, "");
  }

  return s;
}

/**
 * Tính khoảng cách Levenshtein giữa hai chuỗi.
 */
export function levenshteinDistance(s1: string, s2: string): number {
  const m = s1.length;
  const n = s2.length;
  if (m === 0) return n;
  if (n === 0) return m;

  // Sử dụng 2 dòng để tiết kiệm bộ nhớ O(min(m, n))
  let prevRow = new Array(n + 1);
  let currRow = new Array(n + 1);

  for (let j = 0; j <= n; j++) {
    prevRow[j] = j;
  }

  for (let i = 1; i <= m; i++) {
    currRow[0] = i;
    const char1 = s1.charCodeAt(i - 1);

    for (let j = 1; j <= n; j++) {
      const cost = char1 === s2.charCodeAt(j - 1) ? 0 : 1;
      currRow[j] = Math.min(
        currRow[j - 1] + 1,      // chèn
        prevRow[j] + 1,          // xóa
        prevRow[j - 1] + cost    // thay thế
      );
    }

    // Hoán đổi mảng
    for (let j = 0; j <= n; j++) {
      prevRow[j] = currRow[j];
    }
  }

  return prevRow[n];
}

/**
 * Tính độ tương đồng giữa 2 chuỗi từ 0 (hoàn toàn khác) đến 1 (trùng khớp hoàn toàn).
 * Tích hợp Levenshtein và n-gram / token overlap.
 */
export function calculateSimilarity(s1: string, s2: string): number {
  if (s1 === s2) return 1.0;
  if (!s1 || !s2) return 0.0;

  const maxLen = Math.max(s1.length, s2.length);
  if (maxLen === 0) return 1.0;

  // Tính Levenshtein
  const dist = levenshteinDistance(s1, s2);
  const levSim = 1 - dist / maxLen;

  // Nếu chuỗi này chứa trọn vẹn chuỗi kia
  if (s1.includes(s2) || s2.includes(s1)) {
    const minLen = Math.min(s1.length, s2.length);
    const subSim = minLen / maxLen;
    return Math.max(levSim, 0.8 + subSim * 0.2);
  }

  return Math.max(0, Math.min(1, levSim));
}

/**
 * Từ điển các từ viết tắt / đồng nghĩa phổ biến trong thống kê và kế toán
 */
const SYNONYM_MAP: Record<string, string[]> = {
  mst: ["masothue", "ma_so_thue", "msthue", "so_thue", "taxcode"],
  macoso: ["ma_co_so", "macs", "ma_cs", "id_cs", "idcoso", "stt_cs", "stt"],
  madb: ["ma_diaban", "madiaban", "madb", "diaban", "ma_db", "so_db"],
  maxa: ["ma_xa", "maxa", "xa", "xaphuong", "phuongxa", "ma_phuong", "maphuong"],
  mahuyen: ["ma_huyen", "mahuyen", "huyen", "quanhuyen"],
  matinh: ["ma_tinh", "matinh", "tinh", "tinhthanh"],
  manganh: ["ma_nganh", "manganh", "nganh", "vsic", "ma_vsic", "nganhkt", "c5", "cap5"],
  tencoso: ["ten_co_so", "tencs", "tencoso", "donvi", "tendonvi", "ten_don_vi", "ten_dn", "tendoanhnghiep"],
  doanhthu: ["doanh_thu", "doanhthu", "dt", "doanhso", "tong_doanh_thu", "revenue"],
  laodong: ["lao_dong", "laodong", "ld", "tong_ld", "sonhanvien", "nhanvien", "workers"],
  diachi: ["dia_chi", "diachi", "address", "diadiem", "noi_o"],
  dienthoai: ["dien_thoai", "dienthoai", "sdt", "dt", "phone", "mobile"],
};

/**
 * Tự động tìm kiếm các cặp cột giữa File 1 và File 2 có tên giống nhau hoặc đồng nghĩa >= 85%.
 */
export function autoDetectMatchingColumns(
  cols1: string[],
  cols2: string[]
): { col1: string; col2: string; similarity: number; reason: string }[] {
  const matches: { col1: string; col2: string; similarity: number; reason: string }[] = [];

  for (const c1 of cols1) {
    const clean1 = cleanForComparison(c1);
    if (!clean1) continue;

    let bestCol2 = "";
    let bestSim = 0;
    let matchReason = "";

    for (const c2 of cols2) {
      const clean2 = cleanForComparison(c2);
      if (!clean2) continue;

      // 1. Trùng khớp tuyệt đối sau khi bỏ dấu và ký tự đặc biệt
      if (clean1 === clean2) {
        bestCol2 = c2;
        bestSim = 1.0;
        matchReason = "Trùng tên chính xác (đã chuẩn hóa)";
        break;
      }

      // 2. Tra cứu từ điển đồng nghĩa
      let isSynonym = false;
      for (const [key, aliases] of Object.entries(SYNONYM_MAP)) {
        const allKeyVariants = [key, ...aliases];
        const matchC1 = allKeyVariants.some(v => clean1 === cleanForComparison(v) || clean1.includes(cleanForComparison(v)));
        const matchC2 = allKeyVariants.some(v => clean2 === cleanForComparison(v) || clean2.includes(cleanForComparison(v)));
        if (matchC1 && matchC2) {
          isSynonym = true;
          break;
        }
      }

      if (isSynonym && bestSim < 0.95) {
        bestCol2 = c2;
        bestSim = 0.95;
        matchReason = "Đồng nghĩa thường gặp (MST, Xã, Địa bàn, Ngành...)";
        continue;
      }

      // 3. Đo độ tương đồng chuỗi Levenshtein
      const sim = calculateSimilarity(clean1, clean2);
      if (sim > bestSim && sim >= 0.75) {
        bestSim = sim;
        bestCol2 = c2;
        matchReason = `Tên cột tương đồng ${(sim * 100).toFixed(0)}%`;
      }
    }

    if (bestCol2 && bestSim >= 0.75) {
      matches.push({
        col1: c1,
        col2: bestCol2,
        similarity: bestSim,
        reason: matchReason
      });
    }
  }

  // Sắp xếp độ tương đồng từ cao xuống thấp
  return matches.sort((a, b) => b.similarity - a.similarity);
}

export interface SmartLookupResult {
  mergedData: any[];
  addedColumns: string[];
  stats: {
    totalMain: number;
    exactMatches: number;
    fuzzyMatches: number;
    unmatched: number;
    matchRate: number;
  };
  sampleMatches: {
    mainVal: string;
    incomingVal: string;
    similarity: number;
    isExact: boolean;
    rowIdx: number;
    pulledData: Record<string, any>;
  }[];
  sampleUnmatched: {
    rowIdx: number;
    mainVal: string;
    reason: string;
    bestCandidate?: {
      incomingVal: string;
      similarity: number;
    };
  }[];
}

export interface SmartLookupOptions {
  mainData: any[];
  incomingData: any[];
  mainKeyCol: string;
  incomingKeyCol: string;
  columnsToPull: string[]; // Danh sách các cột cần lấy từ incomingData sang mainData
  similarityThreshold?: number; // 0.80 đến 1.0 (mặc định 0.90 = 90%)
  ignoreDiacritics?: boolean;
  ignoreSpaces?: boolean;
  ignorePunctuation?: boolean;
  ignoreCase?: boolean;
  prefixColName?: string; // Tiền tố gắn vào tên cột mới lấy sang (nếu muốn)
}

/**
 * Thực hiện VLOOKUP Mềm thông minh:
 * - Chuẩn hóa bỏ dấu, bỏ khoảng trắng thừa, bỏ ký tự đặc biệt
 * - Tìm kiếm khớp chính xác trước (O(1) Map Lookup)
 * - Đối với các dòng không khớp chính xác, thực hiện dò Fuzzy Match >= threshold
 * - Thu thập thông tin chẩn đoán: vì sao không khớp, giá trị gần giống nhất là gì để người dùng thấy rõ!
 */
export function executeSmartLookup(options: SmartLookupOptions): SmartLookupResult {
  const {
    mainData,
    incomingData,
    mainKeyCol,
    incomingKeyCol,
    columnsToPull,
    similarityThreshold = 0.90,
    ignoreDiacritics = true,
    ignoreSpaces = true,
    ignorePunctuation = true,
    ignoreCase = true,
    prefixColName = "",
  } = options;

  const cleanOpts: CleanOptions = {
    ignoreDiacritics,
    ignoreSpaces,
    ignorePunctuation,
    ignoreCase,
  };

  // 1. Xây dựng chỉ mục cho incomingData
  // Exact Map: cleanedKey -> row
  const exactMap = new Map<string, { row: any; originalKey: string }>();
  // Danh sách để dò Fuzzy Match: { cleanedKey, originalKey, row }
  const indexedIncoming: { cleanedKey: string; originalKey: string; row: any }[] = [];

  for (let i = 0; i < incomingData.length; i++) {
    const row = incomingData[i];
    const rawVal = row[incomingKeyCol];
    if (rawVal === undefined || rawVal === null) continue;
    const strVal = String(rawVal).trim();
    if (!strVal) continue;

    const cleaned = cleanForComparison(strVal, cleanOpts);
    if (!cleaned) continue;

    if (!exactMap.has(cleaned)) {
      exactMap.set(cleaned, { row, originalKey: strVal });
    }
    indexedIncoming.push({ cleanedKey: cleaned, originalKey: strVal, row });
  }

  // Chuẩn bị tên các cột được thêm vào
  const finalAddedCols: { srcCol: string; destCol: string }[] = [];
  const existingColsSet = new Set(mainData.length > 0 ? Object.keys(mainData[0]) : []);

  for (const c of columnsToPull) {
    let destColName = prefixColName ? `${prefixColName}${c}` : c;
    // Nếu tên cột đã tồn tại trong bảng chính và không có tiền tố, tự động thêm tiền tố tệp 2 để tránh đè mất dữ liệu cũ
    if (existingColsSet.has(destColName) && !prefixColName) {
      destColName = `${c}_File2`;
    }
    finalAddedCols.push({ srcCol: c, destCol: destColName });
  }

  let exactMatches = 0;
  let fuzzyMatches = 0;
  let unmatched = 0;

  const sampleMatches: SmartLookupResult["sampleMatches"] = [];
  const sampleUnmatched: SmartLookupResult["sampleUnmatched"] = [];

  // Mảng kết quả sau khi ghép
  const mergedData = new Array(mainData.length);

  for (let i = 0; i < mainData.length; i++) {
    const mainRow = mainData[i];
    const rawVal = mainRow[mainKeyCol];
    const strVal = rawVal !== undefined && rawVal !== null ? String(rawVal).trim() : "";
    const cleanedMain = cleanForComparison(strVal, cleanOpts);

    if (!cleanedMain) {
      // Giá trị rỗng ở file chính -> không thể khớp
      unmatched++;
      if (sampleUnmatched.length < 15) {
        sampleUnmatched.push({
          rowIdx: i + 1,
          mainVal: strVal || "(Ô trống / Rỗng)",
          reason: `Cột "${mainKeyCol}" ở dòng này không có dữ liệu để đối chiếu`,
        });
      }
      // Giữ nguyên dòng và thêm các cột mới với giá trị rỗng
      const addition: Record<string, any> = {};
      for (const col of finalAddedCols) {
        addition[col.destCol] = "";
      }
      mergedData[i] = { ...mainRow, ...addition };
      continue;
    }

    // 1. Thử khớp chính xác sau khi clean (100%)
    const exactHit = exactMap.get(cleanedMain);
    if (exactHit) {
      exactMatches++;
      const addition: Record<string, any> = {};
      const pulledSample: Record<string, any> = {};
      for (const col of finalAddedCols) {
        const val = exactHit.row[col.srcCol] !== undefined ? exactHit.row[col.srcCol] : "";
        addition[col.destCol] = val;
        pulledSample[col.destCol] = val;
      }
      mergedData[i] = { ...mainRow, ...addition };

      if (sampleMatches.length < 15) {
        sampleMatches.push({
          rowIdx: i + 1,
          mainVal: strVal,
          incomingVal: exactHit.originalKey,
          similarity: 1.0,
          isExact: true,
          pulledData: pulledSample,
        });
      }
      continue;
    }

    // 2. Không khớp 100% -> Thử Fuzzy Match nếu ngưỡng < 1.0
    let bestMatch: { cleanedKey: string; originalKey: string; row: any; sim: number } | null = null;
    let highestSim = 0;

    // Quét tìm ứng viên tốt nhất
    for (let j = 0; j < indexedIncoming.length; j++) {
      const inc = indexedIncoming[j];
      const sim = calculateSimilarity(cleanedMain, inc.cleanedKey);
      if (sim > highestSim) {
        highestSim = sim;
        bestMatch = { ...inc, sim };
        if (sim >= 0.98) break; // Khớp gần như tuyệt đối thì dừng sớm
      }
    }

    if (bestMatch && highestSim >= similarityThreshold) {
      // Khớp Fuzzy thành công (đạt ngưỡng)
      fuzzyMatches++;
      const addition: Record<string, any> = {};
      const pulledSample: Record<string, any> = {};
      for (const col of finalAddedCols) {
        const val = bestMatch.row[col.srcCol] !== undefined ? bestMatch.row[col.srcCol] : "";
        addition[col.destCol] = val;
        pulledSample[col.destCol] = val;
      }
      mergedData[i] = { ...mainRow, ...addition };

      if (sampleMatches.length < 15) {
        sampleMatches.push({
          rowIdx: i + 1,
          mainVal: strVal,
          incomingVal: bestMatch.originalKey,
          similarity: highestSim,
          isExact: false,
          pulledData: pulledSample,
        });
      }
    } else {
      // Không tìm thấy hoặc độ giống nhau dưới ngưỡng
      unmatched++;
      if (sampleUnmatched.length < 20) {
        let reason = "Không tìm thấy mã tương ứng bên Tệp 2";
        if (bestMatch && highestSim > 0.4) {
          reason = `Chỉ tìm thấy mã gần nhất là "${bestMatch.originalKey}" (${(highestSim * 100).toFixed(0)}% tương đồng, dưới ngưỡng ${(similarityThreshold * 100).toFixed(0)}%)`;
        }
        sampleUnmatched.push({
          rowIdx: i + 1,
          mainVal: strVal,
          reason,
          bestCandidate: bestMatch && highestSim > 0.4 ? {
            incomingVal: bestMatch.originalKey,
            similarity: highestSim,
          } : undefined,
        });
      }

      const addition: Record<string, any> = {};
      for (const col of finalAddedCols) {
        addition[col.destCol] = "";
      }
      mergedData[i] = { ...mainRow, ...addition };
    }
  }

  const totalMain = mainData.length;
  const matchRate = totalMain > 0 ? ((exactMatches + fuzzyMatches) / totalMain) * 100 : 0;

  return {
    mergedData,
    addedColumns: finalAddedCols.map(c => c.destCol),
    stats: {
      totalMain,
      exactMatches,
      fuzzyMatches,
      unmatched,
      matchRate: Math.round(matchRate * 10) / 10,
    },
    sampleMatches,
    sampleUnmatched,
  };
}
