// vsic.ts - Hoàn toàn động, không chứa dữ liệu cứng

export let vsicRawData: { [key: string]: string } = {};
export let vsicParentMap: { [key: string]: string } = {};

export const normalizeSectorCode = (code: any): string => {
  if (code === null || code === undefined) return "";
  let clean = String(code).trim();
  if (/^[A-Ua-u]$/.test(clean)) return clean.toUpperCase();
  if (clean.toLowerCase().includes("e")) {
    const num = Number(clean);
    if (!isNaN(num)) clean = num.toLocaleString("en-US", { useGrouping: false });
  }
  clean = clean.replace(/\.0+$/, "");
  let digits = clean.replace(/[^\d]/g, "");
  if (digits.length === 7) digits = "0" + digits;
  if (digits.length > 5) digits = digits.substring(0, 5);
  return digits || clean;
};

export const lookupSectorNameWithFallback = (code: string): { name: string; exactMatched: boolean; level: number } => {
  if (!code) return { name: "", exactMatched: false, level: 0 };
  if (vsicRawData[code]) {
    let level = 5;
    if (code.length === 2) level = 2;
    else if (code.length === 3) level = 3;
    else if (code.length === 4) level = 4;
    else if (/^[A-U]$/.test(code)) level = 1;
    return { name: vsicRawData[code], exactMatched: true, level };
  }
  return { name: "", exactMatched: false, level: 0 };
};

export const getSectorHierarchy = (code: string): { [level: string]: { ma: string, ten: string } } => {
  const hierarchy: any = {};
  if (!code) return hierarchy;
  const level1Code = getParentSectorCode(code) || "";
  const level2Code = code.substring(0, 2);
  const level3Code = code.substring(0, 3);
  const level4Code = code.substring(0, 4);
  const level5Code = code.substring(0, 5);
  if (level1Code && vsicRawData[level1Code]) hierarchy['1'] = { ma: level1Code, ten: vsicRawData[level1Code] };
  if (level2Code && vsicRawData[level2Code]) hierarchy['2'] = { ma: level2Code, ten: vsicRawData[level2Code] };
  if (level3Code && vsicRawData[level3Code]) hierarchy['3'] = { ma: level3Code, ten: vsicRawData[level3Code] };
  if (level4Code && vsicRawData[level4Code]) hierarchy['4'] = { ma: level4Code, ten: vsicRawData[level4Code] };
  if (level5Code && vsicRawData[level5Code]) hierarchy['5'] = { ma: level5Code, ten: vsicRawData[level5Code] };
  return hierarchy;
};

export const getSectorLevel = (code: string): number => {
  if (!code) return 0;
  if (/^[A-Z]$/.test(code)) return 1;
  const len = code.length;
  if (len >= 2 && len <= 5) return len;
  return 0;
};

export const getParentSectorCode = (code: string): string => {
  if (!code) return "";
  const cleanKey = code.trim().toUpperCase();
  if (vsicParentMap[cleanKey]) return vsicParentMap[cleanKey];
  if (cleanKey.length > 2) {
    for (let len = cleanKey.length - 1; len >= 2; len--) {
      const sub = cleanKey.substring(0, len);
      if (vsicParentMap[sub]) return vsicParentMap[sub];
    }
  }
  const l2 = cleanKey.substring(0, 2);
  if (vsicParentMap[l2]) return vsicParentMap[l2];
  return "";
};

const removeAccentsAndPunctuation = (str: string): string => {
  if (!str) return "";
  let clean = str.toLowerCase();
  clean = clean.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  clean = clean.replace(/đ/g, "d");
  clean = clean.replace(/[^\w\s]/g, " ");
  clean = clean.replace(/\s+/g, " ");
  return clean.trim();
};

export const smartSuggestSectorByDescription = (description: string): { ma: string; ten: string; diem: number } | null => {
  if (!description) return null;
  const descClean = removeAccentsAndPunctuation(description);
  const descWords = descClean.split(" ").filter(w => w.length >= 2);
  if (descWords.length === 0) return null;
  let bestMatch = { ma: "", ten: "", diem: 0 };
  for (const [code, name] of Object.entries(vsicRawData)) {
    if (code.length !== 5) continue;
    const nameClean = removeAccentsAndPunctuation(name);
    const nameWords = nameClean.split(" ").filter(w => w.length >= 2);
    if (nameWords.length === 0) continue;
    let matchedWordsCount = 0;
    const matchedSet = new Set<string>();
    nameWords.forEach(w => {
      if (descWords.includes(w) && !matchedSet.has(w)) {
        matchedWordsCount++;
        matchedSet.add(w);
      }
    });
    if (matchedWordsCount > 0) {
      const score = (matchedWordsCount / nameWords.length) * (matchedWordsCount / descWords.length);
      if (score > bestMatch.diem) {
        bestMatch = { ma: code, ten: name, diem: score };
      }
    }
  }
  return bestMatch.ma && bestMatch.diem > 0.02 ? bestMatch : null;
};

export const clearAllSectorsInVSIC = () => {
  for (const key in vsicRawData) delete vsicRawData[key];
  for (const key in vsicParentMap) delete vsicParentMap[key];
};

export const loadSectorsIntoVSIC = (catalog: { [key: string]: string }) => {
  Object.assign(vsicRawData, catalog);
};

export const loadParentsIntoVSIC = (parents: { [key: string]: string }) => {
  Object.assign(vsicParentMap, parents);
};

export const isSummaryRow = (row: any): boolean => {
  if (!row || typeof row !== "object") return false;
  for (const key of Object.keys(row)) {
    const val = row[key];
    if (val !== undefined && val !== null) {
      const str = String(val).trim().toLowerCase();
      if (str === "tổng cộng" || str === "cộng" || str === "tổng" || str === "lũy kế" ||
          str === "grand total" || str === "total" || str.startsWith("tổng cộng ") ||
          str.startsWith("lũy kế ") || str.startsWith("cộng ")) {
        return true;
      }
    }
  }
  return false;
};