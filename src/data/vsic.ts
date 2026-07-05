/**
 * @file This file contains the core data and logic for handling VSIC 2018 codes.
 * It includes the raw data map, hierarchy map, and utility functions to work with them.
 * THIS FILE IS AUTO-GENERATED BASED ON THE LATEST VSIC DATA. DO NOT EDIT MANUALLY.
 */

/**
 * vsicRawData
 * A complete key-value map of all VSIC 2018 codes to their official Vietnamese names.
 * This is the primary lookup dictionary.
 * Key: VSIC code (string)
 * Value: Vietnamese name (string)
 */
export const vsicRawData: { [key: string]: string } = {
  // Cấp 1
  "A": "Nông nghiệp, lâm nghiệp và thủy sản",
  "B": "Khai khoáng",
  "C": "Công nghiệp chế biến, chế tạo",
  "D": "Sản xuất và phân phối điện, khí đốt, nước nóng, hơi nước và điều hòa không khí",
  "E": "Cung cấp nước; hoạt động quản lý và xử lý rác thải, nước thải",
  "F": "Xây dựng",
  "G": "Bán buôn và bán lẻ; sửa chữa ô tô, mô tô, xe máy và xe có động cơ khác",
  "H": "Vận tải kho bãi",
  "I": "Dịch vụ lưu trú và ăn uống",
  "J": "Thông tin và truyền thông",
  "K": "Hoạt động tài chính, ngân hàng và bảo hiểm",
  "L": "Hoạt động kinh doanh bất động sản",
  "M": "Hoạt động chuyên môn, khoa học và công nghệ",
  "N": "Hoạt động hành chính và dịch vụ hỗ trợ",
  "O": "Hoạt động của Đảng Cộng sản, tổ chức chính trị - xã hội, quản lý nhà nước, an ninh quốc phòng; bảo đảm xã hội bắt buộc",
  "P": "Giáo dục và đào tạo",
  "Q": "Y tế và hoạt động trợ giúp xã hội",
  "R": "Nghệ thuật, vui chơi và giải trí",
  "S": "Hoạt động dịch vụ khác",
  "T": "Hoạt động làm thuê các công việc trong các hộ gia đình, sản xuất sản phẩm và dịch vụ tự tiêu dùng của hộ gia đình",
  "U": "Hoạt động của các tổ chức và cơ quan quốc tế",

  // Cấp 2
  "01": "Nông nghiệp và hoạt động dịch vụ có liên quan",
  "02": "Lâm nghiệp và hoạt động dịch vụ có liên quan",
  "03": "Khai thác, nuôi trồng thủy sản",
  "05": "Khai thác than cứng và than non",
  "06": "Khai thác dầu thô và khí tự nhiên",
  "07": "Khai thác quặng kim loại",
  "08": "Khai khoáng khác",
  "09": "Hoạt động dịch vụ hỗ trợ khai khoáng",
  "10": "Sản xuất, chế biến thực phẩm",
  "11": "Sản xuất đồ uống",
  "12": "Sản xuất sản phẩm thuốc lá",
  "13": "Dệt",
  "14": "Sản xuất trang phục",
  "15": "Sản xuất da và các sản phẩm có liên quan",
  "16": "Chế biến gỗ và sản xuất sản phẩm từ gỗ, tre, nứa",
  "17": "Sản xuất giấy và sản phẩm từ giấy",
  "18": "In, sao chép bản ghi các loại",
  "19": "Sản xuất than cốc, sản phẩm dầu mỏ tinh chế",
  "20": "Sản xuất hóa chất và sản phẩm hóa chất",
  "21": "Sản xuất thuốc, hóa dược và dược liệu",
  "22": "Sản xuất sản phẩm từ cao su và plastic",
  "23": "Sản xuất sản phẩm từ chất khoáng phi kim loại khác",
  "24": "Sản xuất kim loại",
  "25": "Sản xuất sản phẩm từ kim loại đúc sẵn (trừ máy móc, thiết bị)",
  "26": "Sản xuất sản phẩm điện tử, máy vi tính và sản phẩm quang học",
  "27": "Sản xuất thiết bị điện",
  "28": "Sản xuất máy móc, thiết bị chưa được phân vào đâu",
  "29": "Sản xuất ô tô và xe có động cơ khác",
  "30": "Sản xuất phương tiện vận tải khác",
  "31": "Sản xuất giường, tủ, bàn, ghế",
  "32": "Công nghiệp chế biến, chế tạo khác",
  "33": "Sửa chữa, bảo dưỡng và lắp đặt máy móc và thiết bị",
  "35": "Sản xuất và phân phối điện, khí đốt, nước nóng, hơi nước",
  "36": "Khai thác, xử lý và cung cấp nước",
  "37": "Thoát nước và xử lý nước thải",
  "38": "Thu gom, xử lý và tiêu hủy rác thải; tái chế phế liệu",
  "39": "Xử lý ô nhiễm và hoạt động quản lý chất thải khác",
  "41": "Xây dựng nhà các loại",
  "42": "Xây dựng công trình kỹ thuật dân dụng",
  "43": "Hoạt động xây dựng chuyên dụng",
  "45": "Bán buôn, bán lẻ, sửa chữa ô tô, mô tô, xe máy",
  "46": "Bán buôn (trừ ô tô, mô tô, xe máy)",
  "47": "Bán lẻ (trừ ô tô, mô tô, xe máy)",
  "49": "Vận tải đường bộ và vận tải đường ống",
  "50": "Vận tải đường thủy",
  "51": "Vận tải hàng không",
  "52": "Kho bãi và các hoạt động hỗ trợ cho vận tải",
  "53": "Bưu chính và chuyển phát",
  "55": "Dịch vụ lưu trú",
  "56": "Dịch vụ ăn uống",
  "58": "Hoạt động xuất bản",
  "59": "Hoạt động điện ảnh, sản xuất chương trình truyền hình, ghi âm",
  "60": "Hoạt động phát thanh, truyền hình",
  "61": "Viễn thông",
  "62": "Lập trình máy vi tính, dịch vụ tư vấn và các hoạt động liên quan",
  "63": "Hoạt động dịch vụ thông tin",
  "64": "Hoạt động dịch vụ tài chính (trừ bảo hiểm và bảo hiểm xã hội)",
  "65": "Bảo hiểm, tái bảo hiểm và bảo hiểm xã hội",
  "66": "Hoạt động tài chính khác",
  "68": "Hoạt động kinh doanh bất động sản",
  "69": "Hoạt động pháp luật, kế toán và kiểm toán",
  "70": "Hoạt động của văn phòng sở tại; hoạt động tư vấn quản lý",
  "71": "Hoạt động kiến trúc; kiểm tra và phân tích kỹ thuật",
  "72": "Nghiên cứu khoa học và phát triển công nghệ",
  "73": "Quảng cáo và nghiên cứu thị trường",
  "74": "Hoạt động chuyên môn, khoa học và công nghệ khác",
  "75": "Hoạt động thú y",
  "77": "Hoạt động cho thuê, cho thuê tài chính và hoạt động cho thuê tài sản phi tài chính khác",
  "78": "Hoạt động dịch vụ lao động và việc làm",
  "79": "Hoạt động của các đại lý du lịch, kinh doanh tua du lịch",
  "80": "Hoạt động bảo vệ và điều tra",
  "81": "Hoạt động dịch vụ vệ sinh nhà cửa, công trình và cảnh quan",
  "82": "Hoạt động hành chính, hỗ trợ văn phòng và các hoạt động hỗ trợ kinh doanh khác",
  "84": "Hoạt động quản lý nhà nước, an ninh quốc phòng",
  "85": "Giáo dục và đào tạo",
  "86": "Hoạt động y tế",
  "87": "Hoạt động chăm sóc tập trung trợ giúp xã hội",
  "88": "Hoạt động trợ giúp xã hội không tập trung",
  "89": "Hoạt động của các hiệp hội, tổ chức khác",
  "90": "Hoạt động sáng tạo, nghệ thuật và giải trí",
  "91": "Hoạt động của thư viện, lưu trữ, bảo tàng",
  "92": "Hoạt động xổ số, cá cược và vui chơi giải trí khác",
  "93": "Hoạt động thể thao, vui chơi và giải trí khác",
  "94": "Hoạt động của các hiệp hội, tổ chức khác",
  "95": "Sửa chữa máy vi tính, đồ dùng cá nhân và gia đình",
  "96": "Hoạt động dịch vụ cá nhân khác",
  "97": "Hoạt động của các hộ gia đình đơn lẻ làm thuê",
  "98": "Hoạt động sản xuất sản phẩm vật chất và dịch vụ tự tiêu dùng",
  "99": "Hoạt động của các tổ chức và cơ quan quốc tế"
};

// ----- Utility Functions (DO NOT EDIT) -----

/**
 * Normalizes a VSIC code by removing non-digit characters (unless it is a single-letter code)
 * and ensuring it is a properly padded string format (allowing comparison with vsicRawData).
 * @param code The raw VSIC code from the data.
 * @returns A cleaned, stringified version of the code, or an empty string if invalid.
 */
export const normalizeSectorCode = (code: any): string => {
  if (code === null || code === undefined) return "";
  let clean = String(code).trim();
  
  // 0. KHỚP TUYỆT ĐỐI VỚI DANH MỤC NẠP VÀO TRƯỚC TIÊN (Không biến đổi)
  if (vsicRawData[clean]) {
    return clean;
  }
  const cleanUpper = clean.toUpperCase();
  if (vsicRawData[cleanUpper]) {
    return cleanUpper;
  }

  // If it's a single letter like A, B, C... (Level 1 sector)
  if (/^[A-Za-z]$/.test(clean)) {
    return clean.toUpperCase();
  }

  // Handle scientific notation if any (e.g. 1.11e+07 or 1.11E7) by converting to a flat string
  if (clean.toLowerCase().includes("e")) {
    const num = Number(clean);
    if (!isNaN(num)) {
      clean = num.toLocaleString("en-US", { useGrouping: false });
    }
  }

  // Remove any trailing float suffix (like .0 or .00)
  clean = clean.replace(/\.0+$/, "");

  // Now, extract all digit characters, bypassing layout formatting like dots, spaces, or dashes
  // (e.g. "01.11.00.10" -> "01110010")
  let digits = clean.replace(/[^\d]/g, "");

  if (digits.length > 0) {
    // Heuristic: If it has exactly 7 digits, it is highly likely an 8-digit product code (mã sản phẩm cấp 8) with a stripped leading zero.
    if (digits.length === 7) {
      digits = "0" + digits;
    }

    const is8DigitCode = digits.length === 8;
    
    // 1. Direct check in dictionary
    if (vsicRawData[digits]) {
      return digits;
    }
    
    // 2. Try padding directly based on length first
    if (digits.length > 0 && digits.length < 5) {
      const paddedOneZero = "0" + digits;
      if (vsicRawData[paddedOneZero] || vsicParentMap[paddedOneZero]) {
        return paddedOneZero;
      }
    }

    // 3. Robust Prefix Matching: If "0" + digits matches the start of any custom/local key (e.g. "01110" starts with "0111"), return "0" + digits!
    if (digits.length > 0 && digits.length < 5) {
      const tryPrefix = "0" + digits;
      const dictKeys = Object.keys(vsicRawData);
      const parentKeys = Object.keys(vsicParentMap);
      const matchesDict = dictKeys.some(k => k.startsWith(tryPrefix));
      const matchesParents = parentKeys.some(k => k.startsWith(tryPrefix));
      if (matchesDict || matchesParents) {
        return tryPrefix;
      }
    }
    
    // Fallback: loop-check padding up to 5 digits for general cases
    if (digits.length > 0 && digits.length < 5) {
      for (let len = digits.length + 1; len <= 5; len++) {
        const padded = digits.padStart(len, "0");
        if (vsicRawData[padded]) {
          return padded;
        }
      }
    }
    
    return digits;
  }

  return clean;
};

/**
 * Looks up a sector name in vsicRawData, falling back to parent codes if the exact code is not found.
 * @param code Normalized VSIC code
 * @returns Resolved industry name, plus hierarchical indication if solved via fallback.
 */
export const lookupSectorNameWithFallback = (code: string): { name: string; exactMatched: boolean; level: number } => {
  if (!code) return { name: "", exactMatched: false, level: 0 };

  // Try exact match first in custom-loaded dict
  if (vsicRawData[code]) {
    let level = 5;
    if (code.length === 2) level = 2;
    else if (code.length === 3) level = 3;
    else if (code.length === 4) level = 4;
    else if (/^[A-Z]$/.test(code)) level = 1;
    return { name: vsicRawData[code], exactMatched: true, level };
  }

  // Try looking up clean uppercase direct translation
  const clean = code.trim().toUpperCase();
  if (vsicRawData[clean]) {
    let level = 5;
    if (clean.length === 2) level = 2;
    else if (clean.length === 3) level = 3;
    else if (clean.length === 4) level = 4;
    else if (/^[A-Z]$/.test(clean)) level = 1;
    return { name: vsicRawData[clean], exactMatched: true, level };
  }

  // Try rolling back parent values dynamically based on matching length in currently active vsicRawData,
  // returning the EXACT name in vsicRawData without appending any custom text like "(Quy nạp...)"
  if (clean.length > 1) {
    for (let len = clean.length - 1; len >= 1; len--) {
      const sub = clean.substring(0, len);
      if (vsicRawData[sub]) {
        return {
          name: vsicRawData[sub],
          exactMatched: false,
          level: sub.length === 1 && /^[A-Z]$/.test(sub) ? 1 : sub.length
        };
      }
    }
  }

  return { name: "", exactMatched: false, level: 0 };
};

/**
 * Gets the hierarchical chain for a given VSIC code.
 * For example, for code '46321', it returns an object with names for levels 1, 2, 3, 4, and 5.
 * @param code The VSIC code (normalized).
 * @returns An object where keys are levels ('1' to '5') and values are objects with { ma, ten }.
 */
export const getSectorHierarchy = (code: string): { [level: string]: { ma: string, ten: string } } => {
  const hierarchy: { [level: string]: { ma: string, ten: string } } = {};
  if (!code || typeof code !== 'string') return hierarchy;

  const level1Code = getParentSectorCode(code) || "";
  const level2Code = code.substring(0, 2);
  const level3Code = code.substring(0, 3);
  const level4Code = code.substring(0, 4);
  const level5Code = code.substring(0, 5);
  
  if (level1Code && vsicRawData[level1Code]) {
    hierarchy['1'] = { ma: level1Code, ten: vsicRawData[level1Code] };
  }
  if (level2Code && vsicRawData[level2Code]) {
    hierarchy['2'] = { ma: level2Code, ten: vsicRawData[level2Code] };
  }
  if (level3Code && vsicRawData[level3Code]) {
    hierarchy['3'] = { ma: level3Code, ten: vsicRawData[level3Code] };
  }
  if (level4Code && vsicRawData[level4Code]) {
    hierarchy['4'] = { ma: level4Code, ten: vsicRawData[level4Code] };
  }
  if (level5Code && vsicRawData[level5Code]) {
    hierarchy['5'] = { ma: level5Code, ten: vsicRawData[level5Code] };
  }

  return hierarchy;
};

/**
 * Determines the specific level of a VSIC code.
 * @param code The VSIC code.
 * @returns The level number (1-5) or 0 if not a standard length.
 */
export const getSectorLevel = (code: string): number => {
  if (!code) return 0;
  if (/^[A-Z]$/.test(code)) return 1;
  const len = code.length;
  if (len >= 2 && len <= 5) {
    return len;
  }
  return 0;
};

/**
 * vsicHierarchyMap
 * A pre-computed map for quick parent->child lookups. Not currently used in main logic but available.
 * Key: Parent VSIC code (string)
 * Value: Array of child VSIC codes (string[])
 */
export const vsicHierarchyMap: { [key: string]: string[] } = {};

/**
 * getParentSectorCode
 * Dynamically finds the parent (Level 1) code for any given VSIC code of any length.
 * @param code A normalized VSIC code (of any length from 1 to 5).
 * @returns The corresponding Level 1 letter code ('A' through 'U') or an empty string.
 */
export const getParentSectorCode = (code: string): string => {
    if (!code) return "";
    let cleanKey = code.trim().toUpperCase();

    // 0. Nếu bản thân nó đã là mã cấp 1 (A-U) thì trả về luôn
    if (/^[A-U]$/.test(cleanKey)) {
      return cleanKey;
    }

    const getFlexibleParent = (c: string): string => {
      const cl = c.trim().toUpperCase();
      if (vsicParentMap[cl]) return vsicParentMap[cl];
      const digits = cl.replace(/[^\d]/g, "");
      if (digits.length > 0) {
        if (digits.startsWith("0")) {
          const stripped = digits.replace(/^0+/, "");
          if (vsicParentMap[stripped]) return vsicParentMap[stripped];
        } else {
          const padded = "0" + digits;
          if (vsicParentMap[padded]) return vsicParentMap[padded];
        }
      }
      return "";
    };

    // 1. Đi lên theo sơ đồ cha con vsicParentMap cho đến khi đạt chữ cái cấp 1
    let current = cleanKey;
    let visited = new Set<string>();
    while (current && !/^[A-U]$/.test(current) && !visited.has(current)) {
      visited.add(current);
      const parent = getFlexibleParent(current);
      if (parent) {
        current = parent.trim().toUpperCase();
      } else {
        break;
      }
    }

    if (/^[A-U]$/.test(current)) {
      return current;
    }

    // 2. Nếu không tìm thấy bằng bản đồ cha con, nhưng mã bắt đầu bằng chữ cái từ A đến U
    if (/^[A-U]/.test(cleanKey)) {
      return cleanKey[0];
    }

    // 3. Trích xuất tiền tố dạng số và đối chiếu theo quy ước VSIC chuẩn của Việt Nam
    const digitsOnly = cleanKey.replace(/[^\d]/g, "");
    if (digitsOnly.length >= 2) {
      const prefix = parseInt(digitsOnly.slice(0, 2), 10);
      if (prefix >= 1 && prefix <= 3) return "A";
      if (prefix >= 5 && prefix <= 9) return "B";
      if (prefix >= 10 && prefix <= 33) return "C";
      if (prefix === 35) return "D";
      if (prefix >= 36 && prefix <= 39) return "E";
      if (prefix >= 41 && prefix <= 43) return "F";
      if (prefix >= 45 && prefix <= 47) return "G";
      if (prefix >= 49 && prefix <= 53) return "H";
      if (prefix >= 55 && prefix <= 56) return "I";
      if (prefix >= 58 && prefix <= 63) return "J";
      if (prefix >= 64 && prefix <= 66) return "K";
      if (prefix === 68) return "L";
      if (prefix >= 69 && prefix <= 75) return "M";
      if (prefix >= 77 && prefix <= 82) return "N";
      if (prefix === 84) return "O";
      if (prefix === 85) return "P";
      if (prefix >= 86 && prefix <= 88) return "Q";
      if (prefix >= 90 && prefix <= 93) return "R";
      if (prefix >= 94 && prefix <= 96) return "S";
      if (prefix >= 97 && prefix <= 98) return "T";
      if (prefix === 99) return "U";
    }

    // 4. Các giải pháp cắt giảm tiền tố dự phòng
    if (cleanKey.length > 1) {
      for (let len = cleanKey.length - 1; len >= 1; len--) {
        const sub = cleanKey.substring(0, len);
        if (vsicParentMap[sub]) {
          let subCurrent = vsicParentMap[sub].toUpperCase();
          let subVisited = new Set<string>();
          while (subCurrent && !/^[A-U]$/.test(subCurrent) && !subVisited.has(subCurrent)) {
            subVisited.add(subCurrent);
            const subParent = vsicParentMap[subCurrent];
            if (subParent) {
              subCurrent = subParent.toUpperCase();
            } else {
              break;
            }
          }
          if (/^[A-U]$/.test(subCurrent)) {
            return subCurrent;
          }
        }
      }
    }

    // Fallback cuối cùng
    return "C";
};

/**
 * Helper to remove Vietnamese diacritics and accents plus punctuations
 */
const removeAccentsAndPunctuation = (str: string): string => {
  if (!str) return "";
  let clean = str.toLowerCase();
  clean = clean.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  clean = clean.replace(/đ/g, "d");
  clean = clean.replace(/[^\w\s]/g, " ");
  clean = clean.replace(/\s+/g, " ");
  return clean.trim();
};

/**
 * Suggests a VSIC code based on a textual description by dynamically calculating
 * the word overlap against all Level 5 sectors loaded in memory.
 * @param description The activity description string.
 * @returns An object with the suggested code, its name, and a confidence score.
 */
export const smartSuggestSectorByDescription = (description: string): { ma: string; ten: string; diem: number } | null => {
  if (!description) return null;
  const descClean = removeAccentsAndPunctuation(description);
  const descWords = descClean.split(" ").filter(w => w.length >= 2);
  if (descWords.length === 0) return null;

  let bestMatch = { ma: "", ten: "", diem: 0 };

  // Scan all sectors in the loaded VSIC data
  for (const [code, name] of Object.entries(vsicRawData)) {
    // Prioritize codes of any length from the standard catalog.
    // (We also give a tiny weight to longer codes of length 4 or 5 as tie-breakers)

    const nameClean = removeAccentsAndPunctuation(name);
    const nameWords = nameClean.split(" ").filter(w => w.length >= 2);
    if (nameWords.length === 0) continue;

    // Count overlapping words
    let matchedWordsCount = 0;
    const matchedSet = new Set<string>();
    
    nameWords.forEach(w => {
      if (descWords.includes(w) && !matchedSet.has(w)) {
        matchedWordsCount++;
        matchedSet.add(w);
      }
    });

    if (matchedWordsCount > 0) {
      // Calculate score based on intersection over union of word tokens
      let score = (matchedWordsCount / nameWords.length) * (matchedWordsCount / descWords.length);
      // Tiny adjustment to favor more specific codes (longer is more specify)
      score += (code.length * 0.0001);

      if (score > bestMatch.diem) {
        bestMatch = {
          ma: code,
          ten: name,
          diem: score,
        };
      }
    }
  }

  return bestMatch.ma && bestMatch.diem > 0.02 ? bestMatch : null;
};

// Tự động khôi phục danh bạ ngành nghề tùy chọn do NSD nạp bổ sung từ localStorage khi tải trang
export const vsicParentMap: { [key: string]: string } = {
  // Nhóm con -> Nhóm cha cấp 1 (A-U)
  "01": "A", "02": "A", "03": "A",
  "05": "B", "06": "B", "07": "B", "08": "B", "09": "B",
  "10": "C", "11": "C", "12": "C", "13": "C", "14": "C", "15": "C", "16": "C", "17": "C", "18": "C", "19": "C", "20": "C", "21": "C", "22": "C", "23": "C", "24": "C", "25": "C", "26": "C", "27": "C", "28": "C", "29": "C", "30": "C", "31": "C", "32": "C", "33": "C",
  "35": "D",
  "36": "E", "37": "E", "38": "E", "39": "E",
  "41": "F", "42": "F", "43": "F",
  "45": "G", "46": "G", "47": "G",
  "49": "H", "50": "H", "51": "H", "52": "H", "53": "H",
  "55": "I", "56": "I",
  "58": "J", "59": "J", "60": "J", "61": "J", "62": "J", "63": "J",
  "64": "K", "65": "K", "66": "K",
  "68": "L",
  "69": "M", "70": "M", "71": "M", "72": "M", "73": "M", "74": "M", "75": "M",
  "77": "N", "78": "N", "79": "N", "80": "N", "81": "N", "82": "N",
  "84": "O",
  "85": "P",
  "86": "Q", "87": "Q", "88": "Q",
  "89": "O",
  "90": "R", "91": "R", "92": "R", "93": "R",
  "94": "S", "95": "S", "96": "S",
  "97": "T", "98": "T",
  "99": "U"
};

export const clearAllSectorsInVSIC = () => {
  for (const key in vsicRawData) {
    delete vsicRawData[key];
  }
};

export const clearAllParentsInVSIC = () => {
  for (const key in vsicParentMap) {
    delete vsicParentMap[key];
  }
};

export const loadSectorsIntoVSIC = (catalog: { [key: string]: string }) => {
  Object.assign(vsicRawData, catalog);
};

export const loadParentsIntoVSIC = (parents: { [key: string]: string }) => {
  Object.assign(vsicParentMap, parents);
};

try {
  if (typeof window !== "undefined" && window.localStorage) {
    const isPure = window.localStorage.getItem("custom_vsic_is_pure") === "true";
    if (isPure) {
      clearAllSectorsInVSIC();
      clearAllParentsInVSIC();
    }
    
    const customDataString = window.localStorage.getItem("custom_vsic_data");
    if (customDataString) {
      const customDict = JSON.parse(customDataString);
      Object.assign(vsicRawData, customDict);
    }

    const customParentsString = window.localStorage.getItem("custom_vsic_parents");
    if (customParentsString) {
      const customParents = JSON.parse(customParentsString);
      Object.assign(vsicParentMap, customParents);
    }
  }
} catch (e) {
  console.warn("Lỗi tự động sáp nhập danh mục ngành tùy chọn:", e);
}

/**
 * isSummaryRow
 * Checks if a row in Excel or processed table is a grand total / summary row
 * by scanning its values for keywords like "Tổng cộng", "Cộng", "Total", "Grand Total".
 * @param row The row object.
 * @returns boolean indicating if the row is a total/summary row.
 */
export const isSummaryRow = (row: any): boolean => {
  if (!row || typeof row !== "object") return false;
  for (const key of Object.keys(row)) {
    const val = row[key];
    if (val !== undefined && val !== null) {
      const str = String(val).trim().toLowerCase();
      if (
        str === "tổng cộng" ||
        str === "cộng" ||
        str === "tổng" ||
        str === "lũy kế" ||
        str === "grand total" ||
        str === "total" ||
        str.startsWith("tổng cộng ") ||
        str.startsWith("lũy kế ") ||
        str.startsWith("cộng ") ||
        str.startsWith("cộng:")
      ) {
        return true;
      }
    }
  }
  return false;
};
