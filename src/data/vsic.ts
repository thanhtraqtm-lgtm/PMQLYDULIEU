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
  
  // If it's a single letter like A, B, C... (Level 1 sector)
  if (/^[A-Ua-u]$/.test(clean)) {
    return clean.toUpperCase();
  }

  // Extract the first contiguous sequence of digits (up to 5 characters)
  // This satisfies: "chỉ lấy 5 số đầu và so sánh"
  const digitMatch = clean.match(/\d+/);
  if (digitMatch) {
    let digits = digitMatch[0];
    // If it has more than 5 digits, truncate to first 5 digits
    if (digits.length > 5) {
      digits = digits.substring(0, 5);
    }
    
    // Direct check in dictionary
    if (vsicRawData[digits]) {
      return digits;
    }
    
    // Try padding with leading zeroes if the length is between 1 and 4
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
  
  // Try exact match first
  if (vsicRawData[code]) {
    let level = 5;
    if (code.length === 2) level = 2;
    else if (code.length === 3) level = 3;
    else if (code.length === 4) level = 4;
    else if (/^[A-U]$/.test(code)) level = 1;
    return { name: vsicRawData[code], exactMatched: true, level };
  }

  // If code is numeric, try falling back to parent level codes
  if (/^\d+$/.test(code)) {
    // 5-digit code -> try 4-digit
    if (code.length === 5) {
      const code4 = code.substring(0, 4);
      if (vsicRawData[code4]) {
        return { 
          name: `${vsicRawData[code4]} (Quy nạp từ mã cấp 5: ${code})`, 
          exactMatched: false, 
          level: 4 
        };
      }
    }
    // Try 3-digit
    if (code.length >= 4) {
      const code3 = code.substring(0, 3);
      if (vsicRawData[code3]) {
        return { 
          name: `${vsicRawData[code3]} (Quy nạp từ mã: ${code})`, 
          exactMatched: false, 
          level: 3 
        };
      }
    }
    // Try 2-digit
    if (code.length >= 3) {
      const code2 = code.substring(0, 2);
      if (vsicRawData[code2]) {
        return { 
          name: `${vsicRawData[code2]} (Quy nạp từ mã: ${code})`, 
          exactMatched: false, 
          level: 2 
        };
      }
    }
    // Try Level 1 Letter
    const l2Code = code.substring(0, 2);
    const l1Code = getParentSectorCode(l2Code);
    if (l1Code && vsicRawData[l1Code]) {
      return { 
        name: `${vsicRawData[l1Code]} (Quy nạp từ mã: ${code})`, 
        exactMatched: false, 
        level: 1 
      };
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

  const level1Code = getParentSectorCode(code.substring(0, 2)) || "";
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
 * Dynamically finds the parent (Level 1) code for any given Level 2 code.
 * @param level2Code A two-digit VSIC code.
 * @returns The corresponding Level 1 letter code ('A' through 'U') or an empty string.
 */
export const getParentSectorCode = (level2Code: string): string => {
    if (!level2Code || level2Code.length !== 2) return "";
    const numCode = parseInt(level2Code, 10);
    if (isNaN(numCode)) return "";
  
    if (numCode >= 1 && numCode <= 3) return 'A';
    if (numCode >= 5 && numCode <= 9) return 'B';
    if (numCode >= 10 && numCode <= 33) return 'C';
    if (numCode === 35) return 'D';
    if (numCode >= 36 && numCode <= 39) return 'E';
    if (numCode >= 41 && numCode <= 43) return 'F';
    if (numCode >= 45 && numCode <= 47) return 'G';
    if (numCode >= 49 && numCode <= 53) return 'H';
    if (numCode >= 55 && numCode <= 56) return 'I';
    if (numCode >= 58 && numCode <= 63) return 'J';
    if (numCode >= 64 && numCode <= 66) return 'K';
    if (numCode === 68) return 'L';
    if (numCode >= 69 && numCode <= 75) return 'M';
    if (numCode >= 77 && numCode <= 82) return 'N';
    if (numCode === 84) return 'O';
    if (numCode === 85) return 'P';
    if (numCode >= 86 && numCode <= 88) return 'Q';
    if (numCode >= 90 && numCode <= 93) return 'R';
    if (numCode >= 94 && numCode <= 96) return 'S';
    if (numCode >= 97 && numCode <= 98) return 'T';
    if (numCode === 99) return 'U';
  
    return "";
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
export const clearAllSectorsInVSIC = () => {
  for (const key in vsicRawData) {
    delete vsicRawData[key];
  }
};

export const loadSectorsIntoVSIC = (catalog: { [key: string]: string }) => {
  Object.assign(vsicRawData, catalog);
};

try {
  if (typeof window !== "undefined" && window.localStorage) {
    const isPure = window.localStorage.getItem("custom_vsic_is_pure") === "true";
    if (isPure) {
      clearAllSectorsInVSIC();
    }
    
    const customDataString = window.localStorage.getItem("custom_vsic_data");
    if (customDataString) {
      const customDict = JSON.parse(customDataString);
      Object.assign(vsicRawData, customDict);
    }
  }
} catch (e) {
  console.warn("Lỗi tự động sáp nhập danh mục ngành tùy chọn:", e);
}
