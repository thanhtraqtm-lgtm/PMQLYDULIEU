// Module chuyên sâu xử lý ghép nối, đối soát danh sách hộ khi sáp nhập Tỉnh, Xã, Thôn
// Phục vụ bài toán: Danh sách hộ Điều tra Nông nghiệp (cũ) ↔ Địa bàn / Cơ sở Điều tra Cá thể (mới sau sáp nhập)

import { cleanForComparison, removeVietnameseTones, calculateSimilarity } from "./fuzzyCompare";

/**
 * Loại bỏ các tiền tố hành chính phổ biến trong tiếng Việt để so sánh phần cốt lõi của địa danh.
 * Ví dụ:
 * "Xã Quảng Phúc" -> "quang phuc"
 * "Thị trấn Như Quỳnh" -> "nhu quynh"
 * "Thôn Đồng Cháy 1" -> "dong chay 1"
 * "Tổ dân phố số 5" -> "5" hoặc "so 5"
 * "Bản Nà Lốc" -> "na loc"
 */
export function cleanAdminUnitName(val: any): string {
  if (val === null || val === undefined) return "";
  let s = String(val).trim().toLowerCase();
  if (!s) return "";

  // Bỏ dấu tiếng Việt trước hoặc chuẩn hóa
  const noTone = removeVietnameseTones(s);

  // Xóa các tiền tố cấp xã/phường/thị trấn
  let cleaned = noTone
    .replace(/^(xa|x\.|phuong|p\.|thi tran|tt\.|thi xa|tx\.|thanh pho|tp\.)\s+/gi, "")
    .trim();

  // Xóa các tiền tố cấp thôn/ấp/bản/tổ dân phố
  cleaned = cleaned
    .replace(/^(thon|th\.|ap|ban|lang|buon|plei|soc|to dan pho|tdp|to dp|to|khu pho|kp|khu|cum|khom|doi|dia ban|db|so)\s+/gi, "")
    .trim();

  // Xóa các ký tự đặc biệt còn lại
  cleaned = cleanForComparison(cleaned, {
    ignoreDiacritics: true,
    ignoreCase: true,
    ignorePunctuation: true,
    ignoreSpaces: true,
  });

  return cleaned;
}

/**
 * Chuẩn hóa tên người / chủ hộ:
 * - Bỏ dấu
 * - Bỏ chữ hoa/thường
 * - Bỏ các danh xưng phụ ("ông", "bà", "anh", "chị", "hộ ông", "hộ bà", "chủ hộ")
 */
export function cleanHouseholdHeadName(val: any): string {
  if (val === null || val === undefined) return "";
  let s = String(val).trim().toLowerCase();
  if (!s) return "";

  const noTone = removeVietnameseTones(s);
  // Loại bỏ danh xưng đầu câu
  let cleaned = noTone
    .replace(/^(ho ong|ho ba|ho chu|ong|ba|anh|chi|chu ho|co so|cs|nha ong|nha ba)\s+/gi, "")
    .trim();

  // Loại bỏ ký tự đặc biệt
  cleaned = cleanForComparison(cleaned, {
    ignoreDiacritics: true,
    ignoreCase: true,
    ignorePunctuation: true,
    ignoreSpaces: true,
  });

  return cleaned;
}

/**
 * Đo độ tương đồng giữa 2 tên thôn hoặc 2 tên xã có tính đến tiền tố và số thứ tự
 */
export function calculateAdminSimilarity(nameA: string, nameB: string): number {
  if (!nameA || !nameB) return 0;
  const cA = cleanAdminUnitName(nameA);
  const cB = cleanAdminUnitName(nameB);

  if (cA === cB && cA.length > 0) return 1.0;
  if (!cA || !cB) return 0;

  // Nếu một bên chứa trọn bên kia (VD: "dong chay" và "thon dong chay 1")
  if (cA.includes(cB) || cB.includes(cA)) {
    const minLen = Math.min(cA.length, cB.length);
    const maxLen = Math.max(cA.length, cB.length);
    return Math.max(0.85, (minLen / maxLen) * 0.95);
  }

  return calculateSimilarity(cA, cB);
}

export interface AdminMergeConfig {
  // Cột File A (Kỳ NN cũ)
  colHouseholdHeadA: string;      // Tên chủ hộ
  colCommuneA?: string;           // Tên hoặc mã Xã cũ
  colCommuneCodeA?: string;       // Mã Xã cũ (nếu có riêng)
  colVillageA?: string;           // Tên Thôn / Địa bàn cũ
  colVillageCodeA?: string;       // Mã Thôn cũ (nếu có riêng)
  colIdentifierA?: string;        // CCCD / CMND / SĐT / Mã định danh (nếu có)
  colHouseholdCodeA?: string;     // Mã hộ / STT hộ cũ

  // Cột File B (Kỳ Cá thể mới sau sáp nhập)
  colHouseholdHeadB: string;      // Tên chủ cơ sở / Tên hộ
  colCommuneB?: string;           // Tên Xã mới
  colCommuneCodeB?: string;       // Mã Xã mới
  colVillageB?: string;           // Tên Thôn / Địa bàn mới
  colVillageCodeB?: string;       // Mã Thôn / Địa bàn mới
  colIdentifierB?: string;        // CCCD / CMND / SĐT (nếu có)

  // Bảng quy đổi sáp nhập xã (người dùng tự định nghĩa hoặc do máy sinh)
  // Key: cleaned CommueA -> Value: { communeNameB, communeCodeB }
  communeCrosswalk?: Record<string, { communeNameB: string; communeCodeB: string }>;

  // Ngưỡng tương đồng tối thiểu cho tên chủ hộ (mặc định 0.88)
  nameSimilarityThreshold?: number;
}

export interface MatchedHouseholdRecord {
  // Thông tin gốc Kỳ 1 (NN cũ)
  indexA: number;
  rowA: any;
  householdHeadA: string;
  communeA: string;
  communeCodeA: string;
  villageA: string;
  villageCodeA: string;
  identifierA: string;

  // Thông tin ghép được từ Kỳ 2 (Cá thể mới)
  indexB?: number;
  rowB?: any;
  householdHeadB?: string;
  communeB?: string;
  communeCodeB?: string;
  villageB?: string;
  villageCodeB?: string;
  identifierB?: string;

  // Đánh giá kết quả
  status: "exact_id" | "high_match" | "medium_match" | "crosswalk_match" | "unmatched";
  matchReason: string;
  confidence: number; // 0 - 100%
  candidateCount?: number;
  otherCandidates?: {
    householdHeadB: string;
    communeB: string;
    villageB: string;
    confidence: number;
    rowB: any;
  }[];
}

export interface CommuneMergeCrosswalkItem {
  communeNameA: string;
  communeCodeA: string;
  communeNameB: string;
  communeCodeB: string;
  matchedCount: number;
  confidence: number;
}

export interface VillageMergeCrosswalkItem {
  communeNameA: string;
  villageNameA: string;
  villageCodeA: string;
  communeNameB: string;
  villageNameB: string;
  villageCodeB: string;
  matchedCount: number;
}

export interface AdminMergeMatchingResult {
  records: MatchedHouseholdRecord[];
  communeCrosswalk: CommuneMergeCrosswalkItem[];
  villageCrosswalk: VillageMergeCrosswalkItem[];
  stats: {
    totalA: number;
    totalB: number;
    matchedCount: number;
    exactIdMatches: number;
    highMatches: number;
    mediumMatches: number;
    unmatchedCount: number;
    matchRate: number;
  };
}

/**
 * Thuật toán đa cấp giải quyết bài toán sáp nhập đơn vị hành chính:
 * Vòng 1: Khớp tuyệt đối theo Mã định danh cá nhân (CCCD / SĐT) nếu có.
 * Vòng 2: Tự động học & Khám phá quan hệ sáp nhập Xã cũ -> Xã mới dựa trên các cặp khớp chắc chắn.
 * Vòng 3: Khớp Hộ dựa trên [Tên chủ hộ + Không gian xã sáp nhập + Tên thôn tương đối].
 * Vòng 4: Bóc tách các hộ nghi ngờ/gần khớp để cán bộ rà soát 1 chạm.
 */
export function executeAdminMergeMatching(
  mainData: any[],
  compareData: any[],
  config: AdminMergeConfig
): AdminMergeMatchingResult {
  const {
    colHouseholdHeadA,
    colCommuneA = "",
    colCommuneCodeA = "",
    colVillageA = "",
    colVillageCodeA = "",
    colIdentifierA = "",
    colHouseholdCodeA = "",

    colHouseholdHeadB,
    colCommuneB = "",
    colCommuneCodeB = "",
    colVillageB = "",
    colVillageCodeB = "",
    colIdentifierB = "",

    communeCrosswalk = {},
    nameSimilarityThreshold = 0.88,
  } = config;

  const totalA = mainData.length;
  const totalB = compareData.length;

  // 1. Lập chỉ mục File B
  // - Map theo CCCD/SĐT (nếu có)
  const mapByIdentifierB = new Map<string, { row: any; index: number }>();
  // - Danh sách tất cả các dòng của B được chuẩn hóa
  interface IndexedRowB {
    row: any;
    index: number;
    headRaw: string;
    headClean: string;
    communeRaw: string;
    communeClean: string;
    communeCode: string;
    villageRaw: string;
    villageClean: string;
    villageCode: string;
    idClean: string;
    isClaimed: boolean;
  }

  const indexedB: IndexedRowB[] = [];

  for (let j = 0; j < compareData.length; j++) {
    const row = compareData[j];
    const headRaw = String(row[colHouseholdHeadB] || "").trim();
    const headClean = cleanHouseholdHeadName(headRaw);

    const communeRaw = colCommuneB ? String(row[colCommuneB] || "").trim() : "";
    const communeClean = cleanAdminUnitName(communeRaw);
    const communeCode = colCommuneCodeB ? String(row[colCommuneCodeB] || "").trim() : "";

    const villageRaw = colVillageB ? String(row[colVillageB] || "").trim() : "";
    const villageClean = cleanAdminUnitName(villageRaw);
    const villageCode = colVillageCodeB ? String(row[colVillageCodeB] || "").trim() : "";

    const idRaw = colIdentifierB ? String(row[colIdentifierB] || "").trim() : "";
    const idClean = cleanForComparison(idRaw, { ignorePunctuation: true, ignoreSpaces: true });

    const item: IndexedRowB = {
      row,
      index: j,
      headRaw,
      headClean,
      communeRaw,
      communeClean,
      communeCode,
      villageRaw,
      villageClean,
      villageCode,
      idClean,
      isClaimed: false,
    };

    if (idClean && idClean.length >= 8) {
      if (!mapByIdentifierB.has(idClean)) {
        mapByIdentifierB.set(idClean, { row, index: j });
      }
    }

    indexedB.push(item);
  }

  // Khởi tạo mảng kết quả cho toàn bộ A
  const results: MatchedHouseholdRecord[] = new Array(totalA);

  // Thống kê đếm quan hệ sáp nhập giữa Xã A và Xã B
  // key: `${cleanCommuneA}::${cleanCommuneB}` -> count
  const communeCoOccurrence = new Map<
    string,
    {
      communeNameA: string;
      communeCodeA: string;
      communeNameB: string;
      communeCodeB: string;
      count: number;
    }
  >();

  // key: `${cleanCommuneA}::${cleanVillageA}::${cleanVillageB}` -> count
  const villageCoOccurrence = new Map<
    string,
    {
      communeNameA: string;
      villageNameA: string;
      villageCodeA: string;
      communeNameB: string;
      villageNameB: string;
      villageCodeB: string;
      count: number;
    }
  >();

  // -------------------------------------------------------------------------
  // VÒNG 1: KHỚP THEO CCCD / SĐT / MÃ ĐỊNH DANH (Nếu cả 2 file đều có)
  // -------------------------------------------------------------------------
  let exactIdMatches = 0;

  for (let i = 0; i < totalA; i++) {
    const rowA = mainData[i];
    const headA = String(rowA[colHouseholdHeadA] || "").trim();
    const commA = colCommuneA ? String(rowA[colCommuneA] || "").trim() : "";
    const commCodeA = colCommuneCodeA ? String(rowA[colCommuneCodeA] || "").trim() : "";
    const villA = colVillageA ? String(rowA[colVillageA] || "").trim() : "";
    const villCodeA = colVillageCodeA ? String(rowA[colVillageCodeA] || "").trim() : "";
    const idA = colIdentifierA ? String(rowA[colIdentifierA] || "").trim() : "";
    const idCleanA = cleanForComparison(idA, { ignorePunctuation: true, ignoreSpaces: true });

    if (idCleanA && idCleanA.length >= 8 && mapByIdentifierB.has(idCleanA)) {
      const match = mapByIdentifierB.get(idCleanA)!;
      const rowB = match.row;
      const bIdx = match.index;
      indexedB[bIdx].isClaimed = true;

      const commB = colCommuneB ? String(rowB[colCommuneB] || "").trim() : "";
      const commCodeB = colCommuneCodeB ? String(rowB[colCommuneCodeB] || "").trim() : "";
      const villB = colVillageB ? String(rowB[colVillageB] || "").trim() : "";
      const villCodeB = colVillageCodeB ? String(rowB[colVillageCodeB] || "").trim() : "";
      const headB = String(rowB[colHouseholdHeadB] || "").trim();

      results[i] = {
        indexA: i,
        rowA,
        householdHeadA: headA,
        communeA: commA,
        communeCodeA: commCodeA,
        villageA: villA,
        villageCodeA: villCodeA,
        identifierA: idA,
        indexB: bIdx,
        rowB,
        householdHeadB: headB,
        communeB: commB,
        communeCodeB: commCodeB,
        villageB: villB,
        villageCodeB: villCodeB,
        identifierB: colIdentifierB ? String(rowB[colIdentifierB] || "") : "",
        status: "exact_id",
        matchReason: `Khớp tuyệt đối theo Mã định danh/CCCD/SĐT (${idA})`,
        confidence: 100,
      };
      exactIdMatches++;

      // Ghi nhận vết sáp nhập xã & thôn
      if (commA && commB) {
        const key = `${cleanAdminUnitName(commA)}::${cleanAdminUnitName(commB)}`;
        const existing = communeCoOccurrence.get(key) || {
          communeNameA: commA,
          communeCodeA: commCodeA,
          communeNameB: commB,
          communeCodeB: commCodeB,
          count: 0,
        };
        existing.count++;
        communeCoOccurrence.set(key, existing);
      }

      if (villA && villB) {
        const vKey = `${cleanAdminUnitName(commA)}::${cleanAdminUnitName(villA)}::${cleanAdminUnitName(villB)}`;
        const existingV = villageCoOccurrence.get(vKey) || {
          communeNameA: commA,
          villageNameA: villA,
          villageCodeA: villCodeA,
          communeNameB: commB,
          villageNameB: villB,
          villageCodeB: villCodeB,
          count: 0,
        };
        existingV.count++;
        villageCoOccurrence.set(vKey, existingV);
      }
    }
  }

  // -------------------------------------------------------------------------
  // VÒNG 2: XÂY DỰNG INDEX TÊN CHỦ HỘ VÀ KHÁM PHÁ QUY LUẬT SÁP NHẬP
  // -------------------------------------------------------------------------
  // Tạo Bucket cho File B theo tên chủ hộ sạch để tìm kiếm O(1)
  const bucketByCleanHeadB = new Map<string, IndexedRowB[]>();
  for (let j = 0; j < indexedB.length; j++) {
    const item = indexedB[j];
    if (!item.headClean) continue;
    const list = bucketByCleanHeadB.get(item.headClean) || [];
    list.push(item);
    bucketByCleanHeadB.set(item.headClean, list);
  }

  // Khám phá quan hệ xã sáp nhập từ các hộ trùng tên và thôn tương đồng
  for (let i = 0; i < totalA; i++) {
    if (results[i]) continue; // Đã khớp ở vòng 1
    const rowA = mainData[i];
    const headA = String(rowA[colHouseholdHeadA] || "").trim();
    const headCleanA = cleanHouseholdHeadName(headA);
    if (!headCleanA) continue;

    const commA = colCommuneA ? String(rowA[colCommuneA] || "").trim() : "";
    const commCodeA = colCommuneCodeA ? String(rowA[colCommuneCodeA] || "").trim() : "";
    const villA = colVillageA ? String(rowA[colVillageA] || "").trim() : "";

    const candidates = bucketByCleanHeadB.get(headCleanA);
    if (candidates && candidates.length === 1) {
      // Chỉ có duy nhất 1 người trùng tên trên toàn bộ file B
      const cand = candidates[0];
      const villSim = calculateAdminSimilarity(villA, cand.villageRaw);
      // Nếu thôn trùng hoặc tương đồng >= 70%, đây là cặp đáng tin cậy cao
      if (villSim >= 0.7) {
        const commKey = `${cleanAdminUnitName(commA)}::${cand.communeClean}`;
        const existing = communeCoOccurrence.get(commKey) || {
          communeNameA: commA,
          communeCodeA: commCodeA,
          communeNameB: cand.communeRaw,
          communeCodeB: cand.communeCode,
          count: 0,
        };
        existing.count++;
        communeCoOccurrence.set(commKey, existing);
      }
    }
  }

  // Lập Bảng Ánh Xạ Xã Sáp Nhập Tối Ưu (Best Commune Mapping)
  // Xã cũ A -> Xã mới B có số lượng hộ chuyển sang nhiều nhất
  const autoCommuneMapping = new Map<string, { communeNameB: string; communeCodeB: string; count: number }>();
  for (const [_, info] of communeCoOccurrence.entries()) {
    const cleanA = cleanAdminUnitName(info.communeNameA);
    const currBest = autoCommuneMapping.get(cleanA);
    if (!currBest || info.count > currBest.count) {
      autoCommuneMapping.set(cleanA, {
        communeNameB: info.communeNameB,
        communeCodeB: info.communeCodeB,
        count: info.count,
      });
    }
  }

  // Hòa trộn với bảng quy đổi do người dùng cấu hình thủ công (nếu có)
  for (const [cleanA, customTarget] of Object.entries(communeCrosswalk)) {
    autoCommuneMapping.set(cleanA, {
      communeNameB: customTarget.communeNameB,
      communeCodeB: customTarget.communeCodeB,
      count: 9999, // Ưu tiên quy tắc người dùng
    });
  }

  // -------------------------------------------------------------------------
  // VÒNG 3: GHÉP NỐI TOÀN DIỆN CHO TỪNG HỘ KỲ 1 DỰA TRÊN ÁNH XẠ SÁP NHẬP
  // -------------------------------------------------------------------------
  let highMatches = 0;
  let mediumMatches = 0;

  for (let i = 0; i < totalA; i++) {
    if (results[i]) continue; // Đã khớp ở vòng 1

    const rowA = mainData[i];
    const headA = String(rowA[colHouseholdHeadA] || "").trim();
    const headCleanA = cleanHouseholdHeadName(headA);
    const commA = colCommuneA ? String(rowA[colCommuneA] || "").trim() : "";
    const commCodeA = colCommuneCodeA ? String(rowA[colCommuneCodeA] || "").trim() : "";
    const villA = colVillageA ? String(rowA[colVillageA] || "").trim() : "";
    const villCodeA = colVillageCodeA ? String(rowA[colVillageCodeA] || "").trim() : "";
    const idA = colIdentifierA ? String(rowA[colIdentifierA] || "").trim() : "";

    if (!headCleanA) {
      results[i] = {
        indexA: i,
        rowA,
        householdHeadA: headA,
        communeA: commA,
        communeCodeA: commCodeA,
        villageA: villA,
        villageCodeA: villCodeA,
        identifierA: idA,
        status: "unmatched",
        matchReason: "Tên chủ hộ ở Kỳ 1 bị trống",
        confidence: 0,
      };
      continue;
    }

    const cleanCommA = cleanAdminUnitName(commA);
    const mappedCommuneInfo = autoCommuneMapping.get(cleanCommA);
    const expectedCommuneCleanB = mappedCommuneInfo ? cleanAdminUnitName(mappedCommuneInfo.communeNameB) : "";

    // Tìm kiếm các ứng viên trong File B
    // Chiến lược 1: Tìm theo trùng tên chính xác (Bucket)
    let candidatePool: IndexedRowB[] = bucketByCleanHeadB.get(headCleanA) || [];

    // Nếu đã biết trước Xã mới qua bảng sáp nhập và có nhiều người cùng tên ở nhiều xã khác nhau,
    // ưu tiên lọc đúng người ở Xã mới đó trước!
    if (expectedCommuneCleanB && candidatePool.length > 1) {
      const filteredByCommune = candidatePool.filter(c => c.communeClean === expectedCommuneCleanB);
      if (filteredByCommune.length > 0) {
        candidatePool = filteredByCommune;
      }
    }

    // Chiến lược 2: Nếu không có ai trùng tên chính xác, tìm theo Fuzzy Sim >= nameSimilarityThreshold
    if (candidatePool.length === 0) {
      // Dò trong danh sách B (ưu tiên những dòng thuộc xã dự đoán sáp nhập nếu có)
      for (let j = 0; j < indexedB.length; j++) {
        const itemB = indexedB[j];
        if (expectedCommuneCleanB && itemB.communeClean !== expectedCommuneCleanB) {
          // Bỏ qua nếu không cùng xã sáp nhập để tăng tốc và giảm sai lệch
          continue;
        }
        const sim = calculateSimilarity(headCleanA, itemB.headClean);
        if (sim >= nameSimilarityThreshold) {
          candidatePool.push(itemB);
        }
      }
    }

    // Nếu vẫn không có và có xã sáp nhập dự kiến, mở rộng tìm kiếm với ngưỡng mềm hơn
    if (candidatePool.length === 0 && expectedCommuneCleanB) {
      for (let j = 0; j < indexedB.length; j++) {
        const itemB = indexedB[j];
        if (itemB.communeClean === expectedCommuneCleanB) {
          const sim = calculateSimilarity(headCleanA, itemB.headClean);
          if (sim >= 0.82) {
            candidatePool.push(itemB);
          }
        }
      }
    }

    // Chấm điểm từng ứng viên trong candidatePool
    interface ScoredCandidate {
      itemB: IndexedRowB;
      nameSim: number;
      villSim: number;
      commSim: number;
      isMappedCommune: boolean;
      totalScore: number; // 0 - 100
      reason: string;
    }

    const scoredList: ScoredCandidate[] = [];

    for (const cand of candidatePool) {
      const nameSim = calculateSimilarity(headCleanA, cand.headClean);
      const villSim = calculateAdminSimilarity(villA, cand.villageRaw);
      const isMappedComm = expectedCommuneCleanB ? cand.communeClean === expectedCommuneCleanB : false;
      const commSim = calculateAdminSimilarity(commA, cand.communeRaw);

      // Công thức tính điểm tổng hợp:
      // - Tên chủ hộ: trọng số 55%
      // - Thôn/Địa bàn: trọng số 30%
      // - Xã (khớp qua bảng sáp nhập hoặc tên tương đồng): trọng số 15%
      let score = nameSim * 55;
      if (villSim >= 0.8) {
        score += villSim * 30;
      } else if (villSim >= 0.5) {
        score += villSim * 20;
      } else {
        score += 5; // Có chênh lệch thôn (do sáp nhập thôn)
      }

      if (isMappedComm) {
        score += 15; // Rơi đúng vào xã sáp nhập đã xác định
      } else if (commSim >= 0.8) {
        score += commSim * 12;
      }

      // Điểm cộng nếu chưa có ai nhận dòng này
      if (!cand.isClaimed) {
        score += 3;
      }

      const totalScore = Math.min(100, Math.round(score));

      let reason = `Tên trùng khớp ${(nameSim * 100).toFixed(0)}%`;
      if (isMappedComm) {
        reason += ` • Đúng Xã sáp nhập (${cand.communeRaw})`;
      }
      if (villSim >= 0.8) {
        reason += ` • Khớp Thôn (${cand.villageRaw})`;
      } else if (villA && cand.villageRaw) {
        reason += ` • Thôn cũ: ${villA} ↔ Thôn mới: ${cand.villageRaw}`;
      }

      scoredList.push({
        itemB: cand,
        nameSim,
        villSim,
        commSim,
        isMappedCommune: isMappedComm,
        totalScore,
        reason,
      });
    }

    // Sắp xếp điểm số từ cao xuống thấp
    scoredList.sort((a, b) => b.totalScore - a.totalScore);

    if (scoredList.length > 0 && scoredList[0].totalScore >= 70) {
      const best = scoredList[0];
      const bestItem = best.itemB;
      bestItem.isClaimed = true;

      const isHigh = best.totalScore >= 85;
      if (isHigh) highMatches++;
      else mediumMatches++;

      // Ghi nhận vết sáp nhập xã & thôn
      if (commA && bestItem.communeRaw) {
        const key = `${cleanAdminUnitName(commA)}::${bestItem.communeClean}`;
        const existing = communeCoOccurrence.get(key) || {
          communeNameA: commA,
          communeCodeA: commCodeA,
          communeNameB: bestItem.communeRaw,
          communeCodeB: bestItem.communeCode,
          count: 0,
        };
        existing.count++;
        communeCoOccurrence.set(key, existing);
      }

      if (villA && bestItem.villageRaw) {
        const vKey = `${cleanAdminUnitName(commA)}::${cleanAdminUnitName(villA)}::${bestItem.villageClean}`;
        const existingV = villageCoOccurrence.get(vKey) || {
          communeNameA: commA,
          villageNameA: villA,
          villageCodeA: villCodeA,
          communeNameB: bestItem.communeRaw,
          villageNameB: bestItem.villageRaw,
          villageCodeB: bestItem.villageCode,
          count: 0,
        };
        existingV.count++;
        villageCoOccurrence.set(vKey, existingV);
      }

      results[i] = {
        indexA: i,
        rowA,
        householdHeadA: headA,
        communeA: commA,
        communeCodeA: commCodeA,
        villageA: villA,
        villageCodeA: villCodeA,
        identifierA: idA,
        indexB: bestItem.index,
        rowB: bestItem.row,
        householdHeadB: bestItem.headRaw,
        communeB: bestItem.communeRaw,
        communeCodeB: bestItem.communeCode,
        villageB: bestItem.villageRaw,
        villageCodeB: bestItem.villageCode,
        identifierB: colIdentifierB ? String(bestItem.row[colIdentifierB] || "") : "",
        status: isHigh ? "high_match" : "medium_match",
        matchReason: best.reason,
        confidence: best.totalScore,
        candidateCount: scoredList.length,
        otherCandidates: scoredList.slice(1, 4).map(c => ({
          householdHeadB: c.itemB.headRaw,
          communeB: c.itemB.communeRaw,
          villageB: c.itemB.villageRaw,
          confidence: c.totalScore,
          rowB: c.itemB.row,
        })),
      };
    } else {
      // Không tìm thấy ứng viên đạt chuẩn
      let failReason = "Không tìm thấy chủ hộ tương đồng trên địa bàn mới";
      if (mappedCommuneInfo) {
        failReason += ` (Đã rà soát trong Xã mới: ${mappedCommuneInfo.communeNameB})`;
      }

      results[i] = {
        indexA: i,
        rowA,
        householdHeadA: headA,
        communeA: commA,
        communeCodeA: commCodeA,
        villageA: villA,
        villageCodeA: villCodeA,
        identifierA: idA,
        communeB: mappedCommuneInfo ? mappedCommuneInfo.communeNameB : undefined,
        communeCodeB: mappedCommuneInfo ? mappedCommuneInfo.communeCodeB : undefined,
        status: "unmatched",
        matchReason: failReason,
        confidence: 0,
        candidateCount: scoredList.length,
        otherCandidates: scoredList.slice(0, 3).map(c => ({
          householdHeadB: c.itemB.headRaw,
          communeB: c.itemB.communeRaw,
          villageB: c.itemB.villageRaw,
          confidence: c.totalScore,
          rowB: c.itemB.row,
        })),
      };
    }
  }

  // -------------------------------------------------------------------------
  // TỔNG HỢP BẢNG QUY ĐỔI SÁP NHẬP (CROSSWALK)
  // -------------------------------------------------------------------------
  const communeCrosswalkList: CommuneMergeCrosswalkItem[] = [];
  for (const [_, info] of communeCoOccurrence.entries()) {
    if (info.count >= 1) {
      communeCrosswalkList.push({
        communeNameA: info.communeNameA,
        communeCodeA: info.communeCodeA,
        communeNameB: info.communeNameB,
        communeCodeB: info.communeCodeB,
        matchedCount: info.count,
        confidence: Math.min(100, Math.round(75 + Math.log2(info.count + 1) * 6)),
      });
    }
  }
  communeCrosswalkList.sort((a, b) => b.matchedCount - a.matchedCount);

  const villageCrosswalkList: VillageMergeCrosswalkItem[] = [];
  for (const [_, info] of villageCoOccurrence.entries()) {
    if (info.count >= 1) {
      villageCrosswalkList.push({
        communeNameA: info.communeNameA,
        villageNameA: info.villageNameA,
        villageCodeA: info.villageCodeA,
        communeNameB: info.communeNameB,
        villageNameB: info.villageNameB,
        villageCodeB: info.villageCodeB,
        matchedCount: info.count,
      });
    }
  }
  villageCrosswalkList.sort((a, b) => b.matchedCount - a.matchedCount);

  const matchedCount = exactIdMatches + highMatches + mediumMatches;
  const unmatchedCount = totalA - matchedCount;
  const matchRate = totalA > 0 ? (matchedCount / totalA) * 100 : 0;

  return {
    records: results,
    communeCrosswalk: communeCrosswalkList,
    villageCrosswalk: villageCrosswalkList,
    stats: {
      totalA,
      totalB,
      matchedCount,
      exactIdMatches,
      highMatches,
      mediumMatches,
      unmatchedCount,
      matchRate,
    },
  };
}
