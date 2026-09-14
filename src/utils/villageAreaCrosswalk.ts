/**
 * Thuật toán đối soát và khớp nối Địa bàn / Thôn cũ ➔ Thôn mới & ĐBĐT
 * Dành cho bài toán hơn 1000 địa bàn sáp nhập
 * Kết hợp:
 * 1. Từ điển 10 trang sáp nhập Xã cũ ➔ Xã mới (officialCommuneCrosswalk)
 * 2. Chuẩn hóa tên Thôn / TDP (bóc tách tiền tố TDP, Thôn, Ấp, số La Mã...)
 * 3. So khớp tổng số hộ của thôn (chìa khóa vàng đối chiếu chéo)
 */

import {
  OFFICIAL_COMMUNE_MERGE_DATA,
  cleanCommuneStr,
  findOfficialNewCommune,
  findOfficialByCommuneCode,
  findOfficialNewCommuneCode,
  findAllOldCommunesForNew,
} from "../data/officialCommuneCrosswalk";
export {
  OFFICIAL_COMMUNE_MERGE_DATA,
  cleanCommuneStr,
  findOfficialNewCommune,
  findOfficialByCommuneCode,
  findOfficialNewCommuneCode,
  findAllOldCommunesForNew,
};

export interface VillageRecordA {
  raw: any;
  indexA: number;
  communeNameA: string;
  communeCodeA: string;
  villageNameA: string;
  villageCodeA: string;
  householdCountA: number;
}

export interface VillageRecordB {
  raw: any;
  indexB: number;
  communeNameB: string;
  communeCodeB: string;
  villageNameB: string;
  villageCodeB: string;
  householdCountB: number;
  dbdtNameB: string;
  dbdtCodeB: string;
  dbdtHouseholdCountB: number;
}

export interface VillageMatchCandidate {
  recordB: VillageRecordB;
  confidence: number;
  matchGrade: "PERFECT" | "VERY_HIGH" | "HIGH" | "MEDIUM" | "LOW";
  matchReason: string;
  nameSimilarity: number;
  diffHouseholds: number;
  diffPercentage: number;
}

export interface VillageMatchResult {
  recordA: VillageRecordA;
  matchedB: VillageRecordB | null;
  confidence: number;
  matchGrade: "PERFECT" | "VERY_HIGH" | "HIGH" | "MEDIUM" | "LOW" | "UNMATCHED";
  matchReason: string;
  nameSimilarity: number;
  diffHouseholds: number;
  diffPercentage: number;
  candidates: VillageMatchCandidate[];
  manualOverride?: boolean;
}

export interface VillageAreaCrosswalkConfig {
  colCommuneNameA: string;
  colCommuneCodeA: string;
  colVillageNameA: string;
  colVillageCodeA: string;
  colHouseholdsA: string;

  colCommuneNameB: string;
  colCommuneCodeB: string;
  colVillageNameB: string;
  colVillageCodeB: string;
  colHouseholdsB: string;
  colDbdtNameB: string;
  colDbdtCodeB: string;
  colDbdtHouseholdsB: string;

  customCommuneCrosswalk?: { communeA: string; communeB: string; communeCodeB?: string }[];
  maxHouseholdDiffPercent?: number; // Mặc định 25%
  selectedCommuneA?: string; // Chọn lọc chạy riêng cho 1 xã cụ thể
  strictCommuneScoped?: boolean; // Khóa chặt phạm vi xã: chỉ so khớp các địa bàn trong cùng xã hoặc cặp sáp nhập tương ứng (Mặc định true)
}

/**
 * Chuẩn hóa tên Thôn / Tổ dân phố:
 * Bóc tách tiền tố "Tổ dân phố", "TDP", "Thôn", "Ấp", "Bản", "Khóm", "Khu phố", "Khu", "Cụm", "Xóm"...
 */
export function cleanVillageName(name: string): {
  original: string;
  cleaned: string;
  noDiacritics: string;
  detectedPrefix: string;
} {
  if (!name || typeof name !== "string") {
    return { original: "", cleaned: "", noDiacritics: "", detectedPrefix: "" };
  }

  let str = name.trim();
  let detectedPrefix = "";

  // Danh mục tiền tố thôn/tổ dân phố
  const prefixPatterns = [
    /^(tổ\s*dân\s*phố|to\s*dan\s*pho|tdp|tổ\s*dp|to\s*dp|t\.\s*d\.\s*p)\s+/i,
    /^(thôn|thon|ấp|ap|bản|ban|khóm|khom|khu\s*phố|khu\s*pho|kp|k\.\s*p)\s+/i,
    /^(khu\s*dân\s*cư|khu\s*dan\s*cu|kdc|khu|cụm|cum|xóm|xom|đội|doi)\s+/i,
    /^(địa\s*bàn\s*điều\s*tra|địa\s*bàn|đbđt|dbdt)\s+/i,
  ];

  for (const pat of prefixPatterns) {
    const m = str.match(pat);
    if (m) {
      detectedPrefix = m[1];
      str = str.replace(pat, "").trim();
      break;
    }
  }

  // Chuẩn hóa khoảng trắng
  const cleaned = str.replace(/\s+/g, " ").toLowerCase();

  // Bỏ dấu tiếng Việt
  const noDiacritics = cleaned
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase();

  return {
    original: name.trim(),
    cleaned,
    noDiacritics,
    detectedPrefix,
  };
}

/**
 * Chuẩn hóa số La Mã sang số thường để so sánh (VD: "Điện Biên II" ➔ "Điện Biên 2")
 */
function normalizeRomanNumerals(str: string): string {
  return str
    .replace(/\bviii\b/gi, "8")
    .replace(/\bvii\b/gi, "7")
    .replace(/\bvi\b/gi, "6")
    .replace(/\biv\b/gi, "4")
    .replace(/\bv\b/gi, "5")
    .replace(/\biii\b/gi, "3")
    .replace(/\bii\b/gi, "2")
    .replace(/\bi\b/gi, "1")
    .replace(/\bix\b/gi, "9")
    .replace(/\bx\b/gi, "10");
}

/**
 * Tính độ tương đồng giữa 2 chuỗi tên thôn
 */
export function computeNameSimilarity(nameA: string, nameB: string): number {
  const normA = cleanVillageName(nameA);
  const normB = cleanVillageName(nameB);

  if (normA.cleaned === normB.cleaned) return 1.0;
  if (normA.noDiacritics === normB.noDiacritics) return 0.98;

  // So sánh sau khi quy đổi số La Mã
  const romanA = normalizeRomanNumerals(normA.noDiacritics);
  const romanB = normalizeRomanNumerals(normB.noDiacritics);
  if (romanA === romanB) return 0.96;

  // Bigram Dice Coefficient
  const sA = romanA.replace(/[^a-z0-9]/g, "");
  const sB = romanB.replace(/[^a-z0-9]/g, "");

  if (!sA || !sB) return 0;
  if (sA === sB) return 0.95;

  // Token word matching
  const wordsA = romanA.split(/\s+/).filter(Boolean);
  const wordsB = romanB.split(/\s+/).filter(Boolean);
  const setB = new Set(wordsB);
  let commonWords = 0;
  for (const w of wordsA) {
    if (setB.has(w)) commonWords++;
  }
  const tokenJaccard = commonWords / Math.max(wordsA.length, wordsB.length, 1);

  const getBigrams = (str: string) => {
    const s = new Set<string>();
    for (let i = 0; i < str.length - 1; i++) {
      s.add(str.slice(i, i + 2));
    }
    return s;
  };

  const bigramsA = getBigrams(sA);
  const bigramsB = getBigrams(sB);
  let intersection = 0;
  bigramsA.forEach(b => {
    if (bigramsB.has(b)) intersection++;
  });
  const dice = (2 * intersection) / (bigramsA.size + bigramsB.size || 1);

  return Math.max(dice * 0.85, tokenJaccard * 0.9);
}

/**
 * Lấy số nguyên an toàn từ giá trị ô Excel (hỗ trợ định dạng có dấu chấm/phẩy)
 */
export function parseHouseholdCount(val: any): number {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return Math.round(val);
  const s = String(val).trim().replace(/\s+/g, "").replace(/,/g, "");
  const n = parseInt(s, 10);
  return isNaN(n) ? 0 : n;
}

// cleanCommuneStr is imported and re-exported from officialCommuneCrosswalk

/**
 * Thuật toán khớp nối chính
 */
export function executeVillageAreaCrosswalk(
  rowsA: any[],
  rowsB: any[],
  config: VillageAreaCrosswalkConfig
): {
  results: VillageMatchResult[];
  summary: {
    totalA: number;
    totalB: number;
    perfectMatches: number;
    veryHighMatches: number;
    highMatches: number;
    mediumMatches: number;
    unmatched: number;
    matchRate: number;
  };
} {
  // 1. Chuẩn hóa dữ liệu Tệp A
  const recordsA: VillageRecordA[] = rowsA.map((r, idx) => ({
    raw: r,
    indexA: idx,
    communeNameA: String(r[config.colCommuneNameA] || "").trim(),
    communeCodeA: String(r[config.colCommuneCodeA] || "").trim(),
    villageNameA: String(r[config.colVillageNameA] || "").trim(),
    villageCodeA: String(r[config.colVillageCodeA] || "").trim(),
    householdCountA: parseHouseholdCount(r[config.colHouseholdsA]),
  }));

  // 2. Chuẩn hóa dữ liệu Tệp B
  const recordsB: VillageRecordB[] = rowsB.map((r, idx) => ({
    raw: r,
    indexB: idx,
    communeNameB: String(r[config.colCommuneNameB] || "").trim(),
    communeCodeB: String(r[config.colCommuneCodeB] || "").trim(),
    villageNameB: String(r[config.colVillageNameB] || "").trim(),
    villageCodeB: String(r[config.colVillageCodeB] || "").trim(),
    householdCountB: parseHouseholdCount(r[config.colHouseholdsB]),
    dbdtNameB: String(r[config.colDbdtNameB] || "").trim(),
    dbdtCodeB: String(r[config.colDbdtCodeB] || "").trim(),
    dbdtHouseholdCountB: parseHouseholdCount(r[config.colDbdtHouseholdsB]),
  }));

  // 3. Xây dựng bản đồ Ánh xạ Xã: Kết hợp 310 xã chính thức + quy tắc tùy biến
  const communeRuleMap = new Map<string, string>();
  const communeCodeRuleMap = new Map<string, { newCode: string; newName: string }>();

  // Nạp 310 xã từ tài liệu chính thức 9 trang của người dùng
  for (const rule of OFFICIAL_COMMUNE_MERGE_DATA) {
    const k = cleanCommuneStr(rule.communeA);
    const v = cleanCommuneStr(rule.communeB);
    if (k && v) communeRuleMap.set(k, v);
    if (rule.codeA && rule.codeB) {
      communeCodeRuleMap.set(String(rule.codeA).trim(), {
        newCode: String(rule.codeB).trim(),
        newName: rule.communeB,
      });
    }
  }

  // Nạp bổ sung quy tắc tùy biến của người dùng nếu có
  if (config.customCommuneCrosswalk) {
    for (const rule of config.customCommuneCrosswalk) {
      if (rule.communeA && rule.communeB) {
        communeRuleMap.set(cleanCommuneStr(rule.communeA), cleanCommuneStr(rule.communeB));
      }
    }
  }

  // Lọc theo xã nếu người dùng chọn chạy riêng 1 xã cụ thể
  let activeRecordsA = recordsA;
  if (config.selectedCommuneA && config.selectedCommuneA !== "ALL") {
    const selClean = cleanCommuneStr(config.selectedCommuneA);
    activeRecordsA = recordsA.filter(r => cleanCommuneStr(r.communeNameA) === selClean);
  }

  const results: VillageMatchResult[] = [];

  for (const recA of activeRecordsA) {
    const hCountA = recA.householdCountA;

    // Xác định xã mới mục tiêu (tra cứu theo cả Mã xã cũ và Tên xã cũ)
    const cleanedCommuneA = cleanCommuneStr(recA.communeNameA);
    let targetNewCommune = communeRuleMap.get(cleanedCommuneA) || null;
    let targetNewCommuneCode: string | null = null;

    if (recA.communeCodeA && communeCodeRuleMap.has(recA.communeCodeA)) {
      const foundRule = communeCodeRuleMap.get(recA.communeCodeA)!;
      targetNewCommuneCode = foundRule.newCode;
      if (!targetNewCommune) targetNewCommune = cleanCommuneStr(foundRule.newName);
    }

    if (!targetNewCommune && recA.communeNameA) {
      const found = findOfficialNewCommune(recA.communeNameA);
      if (found) targetNewCommune = cleanCommuneStr(found);
    }
    if (!targetNewCommuneCode && recA.communeCodeA) {
      const foundCode = findOfficialNewCommuneCode(recA.communeCodeA);
      if (foundCode) targetNewCommuneCode = foundCode;
    }

    const candidates: VillageMatchCandidate[] = [];

    for (const recB of recordsB) {
      const hCountB = recB.householdCountB;
      const cleanedCommuneB = cleanCommuneStr(recB.communeNameB);

      // Kiểm tra tương thích cấp Xã
      let communeStatus: "MATCH" | "MISMATCH" | "UNKNOWN" = "UNKNOWN";

      // 1. Kiểm tra bằng Mã Xã (chuẩn nhất)
      if (targetNewCommuneCode && recB.communeCodeB) {
        if (recB.communeCodeB === targetNewCommuneCode) {
          communeStatus = "MATCH";
        }
      }

      // 2. Kiểm tra bằng Tên Xã Mới mục tiêu
      if (communeStatus !== "MATCH" && targetNewCommune && cleanedCommuneB) {
        if (
          cleanedCommuneB === targetNewCommune ||
          cleanedCommuneB.includes(targetNewCommune) ||
          targetNewCommune.includes(cleanedCommuneB)
        ) {
          communeStatus = "MATCH";
        } else {
          communeStatus = "MISMATCH";
        }
      } else if (communeStatus !== "MATCH" && cleanedCommuneA && cleanedCommuneB) {
        // Cùng tên xã hoặc cùng mã xã (không sáp nhập hoặc cùng đơn vị)
        if (recA.communeCodeA && recB.communeCodeB && recA.communeCodeA === recB.communeCodeB) {
          communeStatus = "MATCH";
        } else if (
          cleanedCommuneA === cleanedCommuneB ||
          cleanedCommuneB.includes(cleanedCommuneA) ||
          cleanedCommuneA.includes(cleanedCommuneB)
        ) {
          communeStatus = "MATCH";
        } else {
          communeStatus = "MISMATCH";
        }
      }

      // CHẾ ĐỘ CHẠY THEO XÃ / ĐỊA BÀN TRONG XÃ VỚI NHAU:
      // Nếu bật strictCommuneScoped (mặc định true): Tuyệt đối không so sánh địa bàn sang xã khác
      if (config.strictCommuneScoped !== false) {
        if (communeStatus === "MISMATCH") {
          continue;
        }
      }

      // 1. Tính độ tương đồng tên thôn
      const nameSim = computeNameSimilarity(recA.villageNameA, recB.villageNameB);

      // Nếu khác xã dự kiến và tên cũng không khớp cao thì bỏ qua ngay để tăng tốc cho 1000+ dòng
      if (communeStatus === "MISMATCH" && nameSim < 0.85) continue;
      if (nameSim < 0.35) continue;

      // 2. Tính chênh lệch số hộ của thôn
      const diffHouseholds = Math.abs(hCountA - hCountB);
      const maxH = Math.max(hCountA, hCountB, 1);
      const diffPercentage = (diffHouseholds / maxH) * 100;

      // 3. Tính điểm tổng hợp & xếp loại
      let score = 0;
      let reason = "";
      let grade: VillageMatchCandidate["matchGrade"] = "LOW";

      if (nameSim >= 0.95 && diffHouseholds === 0 && hCountA > 0) {
        score = 100;
        reason = `Khớp tuyệt đối 100%: Trùng tên thôn & Số hộ bằng nhau chằn chặn (${hCountA} hộ)`;
        grade = "PERFECT";
      } else if (nameSim >= 0.95 && diffHouseholds <= 2 && hCountA > 0) {
        score = 99;
        reason = `Khớp cực cao: Trùng tên thôn & Lệch chỉ ${diffHouseholds} hộ (${hCountA} vs ${hCountB})`;
        grade = "PERFECT";
      } else if (nameSim >= 0.95 && diffPercentage <= 5 && hCountA > 0) {
        score = 96;
        reason = `Khớp rất cao: Trùng tên thôn & Số hộ lệch ${diffHouseholds} hộ (${diffPercentage.toFixed(1)}%)`;
        grade = "VERY_HIGH";
      } else if (nameSim >= 0.95 && diffPercentage <= 15) {
        score = 90;
        reason = `Khớp cao: Trùng tên thôn, số hộ lệch ${diffHouseholds} hộ (${diffPercentage.toFixed(1)}%)`;
        grade = "HIGH";
      } else if (nameSim >= 0.85 && diffPercentage <= 10 && hCountA > 0) {
        score = 88;
        reason = `Tên tương đồng (${Math.round(nameSim * 100)}%) + Số hộ sát nhau (lệch ${diffHouseholds} hộ)`;
        grade = "HIGH";
      } else if (nameSim >= 0.95) {
        score = 75;
        reason = `Trùng tên thôn nhưng số hộ lệch nhiều (${hCountA} vs ${hCountB})`;
        grade = "MEDIUM";
      } else if (nameSim >= 0.70 && diffPercentage <= 5 && hCountA > 0) {
        score = 80;
        reason = `Tên xấp xỉ (${Math.round(nameSim * 100)}%) nhưng số hộ gần như bằng nhau (lệch ${diffHouseholds} hộ)`;
        grade = "HIGH";
      } else if (nameSim >= 0.65) {
        score = 60;
        reason = `Tên tương đồng ${Math.round(nameSim * 100)}% (Cần xác nhận)`;
        grade = "LOW";
      } else {
        score = 45;
        reason = `Độ tin cậy thấp`;
        grade = "LOW";
      }

      // Điểm cộng nếu khớp chính xác Xã mới từ tài liệu 10 trang
      if (communeStatus === "MATCH") {
        score = Math.min(100, score + 5);
        reason += ` • Đã khớp xã mới [${recB.communeNameB}] theo Quyết định sáp nhập`;
      } else if (communeStatus === "MISMATCH") {
        score = Math.max(10, score - 25);
        reason += ` • Cảnh báo: Thuộc xã [${recB.communeNameB}] khác xã dự kiến`;
      }

      if (score >= 45) {
        candidates.push({
          recordB: recB,
          confidence: score,
          matchGrade: grade,
          matchReason: reason,
          nameSimilarity: nameSim,
          diffHouseholds,
          diffPercentage,
        });
      }
    }

    // Sắp xếp ứng viên từ điểm cao nhất
    candidates.sort((a, b) => b.confidence - a.confidence);

    const best = candidates[0] || null;

    if (best && best.confidence >= 65) {
      results.push({
        recordA: recA,
        matchedB: best.recordB,
        confidence: best.confidence,
        matchGrade: best.matchGrade,
        matchReason: best.matchReason,
        nameSimilarity: best.nameSimilarity,
        diffHouseholds: best.diffHouseholds,
        diffPercentage: best.diffPercentage,
        candidates,
      });
    } else {
      results.push({
        recordA: recA,
        matchedB: null,
        confidence: 0,
        matchGrade: "UNMATCHED",
        matchReason: candidates.length > 0 ? "Chưa đủ tin cậy (Cần xác nhận thủ công)" : "Không tìm thấy địa bàn mới tương ứng",
        nameSimilarity: 0,
        diffHouseholds: 0,
        diffPercentage: 0,
        candidates,
      });
    }
  }

  const perfectMatches = results.filter(r => r.matchGrade === "PERFECT").length;
  const veryHighMatches = results.filter(r => r.matchGrade === "VERY_HIGH").length;
  const highMatches = results.filter(r => r.matchGrade === "HIGH").length;
  const mediumMatches = results.filter(r => r.matchGrade === "MEDIUM").length;
  const unmatched = results.filter(r => r.matchGrade === "UNMATCHED").length;
  const matchedTotal = results.length - unmatched;
  const matchRate = results.length > 0 ? Math.round((matchedTotal / results.length) * 100) : 0;

  return {
    results,
    summary: {
      totalA: recordsA.length,
      totalB: recordsB.length,
      perfectMatches,
      veryHighMatches,
      highMatches,
      mediumMatches,
      unmatched,
      matchRate,
    },
  };
}
