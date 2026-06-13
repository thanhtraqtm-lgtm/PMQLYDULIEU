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
export const vsicRawData: { [key: string]: string } = {};

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

    // Truncate to first 5 digits for standard level 5 industry comparison
    if (digits.length > 5) {
      digits = digits.substring(0, 5);
    }
    
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
    
    // 1. Kiểm tra ánh xạ cha con tùy chỉnh trực tiếp tuyệt đối (vd: "01", "011", "01110" v.v.)
    const cleanKey = code.trim().toUpperCase();
    if (vsicParentMap[cleanKey]) {
      return vsicParentMap[cleanKey];
    }

    // 2. Kiểm tra nếu mã cha có chiều dài lớn và có thể cắt nhỏ dần để tìm
    if (cleanKey.length > 1) {
      for (let len = cleanKey.length - 1; len >= 1; len--) {
        const sub = cleanKey.substring(0, len);
        if (vsicParentMap[sub]) {
          return vsicParentMap[sub];
        }
      }
    }

    // 3. Nếu không có ánh xạ, xác định mã cha theo quy ước tiền tố xuất hiện sớm nhất trong Danh mục nạp của người dùng
    if (/^[A-Z]/.test(cleanKey)) {
      return cleanKey[0];
    }

    const keys = Object.keys(vsicRawData);
    let shortestParent = "";
    keys.forEach(k => {
      if (cleanKey.startsWith(k) && k !== cleanKey) {
        if (!shortestParent || k.length < shortestParent.length) {
          shortestParent = k;
        }
      }
    });

    if (shortestParent) {
      return shortestParent;
    }

    // Fallback: Lấy 2 ký số đầu cho nhóm cấp 2 tiêu chuẩn hoặc toàn bộ nếu ngắn dưới 2 ký tự
    if (cleanKey.length >= 2) {
      return cleanKey.substring(0, 2);
    }
  
    return cleanKey;
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
    // We target Level 5 codes (typically 5 digits) for specific mapping suggestions
    if (code.length !== 5) continue;

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
      const score = (matchedWordsCount / nameWords.length) * (matchedWordsCount / descWords.length);
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
export const vsicParentMap: { [key: string]: string } = {};

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
    const isPure = window.localStorage.getItem("custom_vsic_is_pure") !== "false";
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
