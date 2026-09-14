import { OFFICIAL_COMMUNE_MERGE_DATA, cleanCommuneStr, findOfficialNewCommune } from "./villageAreaCrosswalk";

/**
 * Thuật toán Đối Soát Cụm Hộ Dân (Household Cluster Overlap)
 * Dùng khi: Đã so đến cấp thôn nhưng vẫn chưa tìm thấy địa bàn do tên thôn bị đổi hoàn toàn,
 * sáp nhập nhiều thôn hoặc mất mã.
 * 
 * Nguyên lý:
 * Gom toàn bộ danh sách hộ/chủ hộ của từng Địa bàn cũ (Kỳ A) và từng Địa bàn mới (Kỳ B).
 * Quét đối chiếu ma trận chéo họ tên người dân.
 * Nếu phát hiện 2 địa bàn có từ N người trùng tên trở lên (mặc định >= 30 người)
 * ➔ Kết luận tạm tính: 2 địa bàn này chính là một (hoặc đã sáp nhập vào nhau).
 */

export interface RawHouseholdItem {
  raw: any;
  rowIndex: number;
  fullName: string;
  normalizedName: string;
  noDiacriticsName: string;
  areaKey: string;
  communeName: string;
  communeCode: string;
  villageName: string;
  villageCode: string;
  dbdtName?: string;
  dbdtCode?: string;
  extraInfo?: string; // Năm sinh, CCCD, SĐT hoặc địa chỉ nếu có
}

export interface AreaCluster {
  areaKey: string;
  communeName: string;
  communeCode: string;
  villageName: string;
  villageCode: string;
  dbdtName: string;
  dbdtCode: string;
  households: RawHouseholdItem[];
  totalHouseholds: number;
}

export interface HouseholdMatchedDetail {
  nameA: string;
  nameB: string;
  rowA: number;
  rowB: number;
  extraA?: string;
  extraB?: string;
  similarity: number;
}

export interface AreaClusterCandidate {
  clusterB: AreaCluster;
  overlapCount: number;
  overlapRatioA: number; // Tỷ lệ trên tổng hộ của Thôn A
  overlapRatioB: number; // Tỷ lệ trên tổng hộ của Thôn B
  confidence: number;
  matchGrade: "PERFECT_CLUSTER" | "HIGH_CLUSTER" | "MEDIUM_CLUSTER" | "LOW_CLUSTER";
  matchReason: string;
  matchedHouseholds: HouseholdMatchedDetail[];
}

export interface AreaClusterMatchResult {
  clusterA: AreaCluster;
  matchedClusterB: AreaCluster | null;
  overlapCount: number;
  overlapRatioA: number;
  overlapRatioB: number;
  confidence: number;
  matchGrade: "PERFECT_CLUSTER" | "HIGH_CLUSTER" | "MEDIUM_CLUSTER" | "UNMATCHED";
  matchReason: string;
  matchedHouseholds: HouseholdMatchedDetail[];
  candidates: AreaClusterCandidate[];
  manualApproved?: boolean;
}

export interface HouseholdClusterConfig {
  colPersonNameA: string;
  colAreaOrVillageA: string;
  colCommuneA?: string;
  colVillageCodeA?: string;
  colExtraA?: string; // CCCD, năm sinh, địa chỉ...

  colPersonNameB: string;
  colAreaOrVillageB: string;
  colCommuneB?: string;
  colVillageCodeB?: string;
  colDbdtNameB?: string;
  colDbdtCodeB?: string;
  colExtraB?: string;

  minOverlapThreshold: number; // Mặc định 30 người trùng tên
  minOverlapRatio?: number; // Mặc định 25% (nếu thôn nhỏ dưới 30 người)
  useFuzzyName?: boolean; // Cho phép sai lệch nhẹ 1 ký tự
  selectedCommuneA?: string; // Chọn lọc chạy riêng cho 1 xã cụ thể
  strictCommuneScoped?: boolean; // Khóa chặt phạm vi xã: chỉ so khớp các địa bàn trong cùng xã hoặc cặp sáp nhập (Mặc định true)
}

/**
 * Chuẩn hóa họ tên người dân: bỏ khoảng cách thừa, bỏ ký tự đặc biệt, chuẩn hóa tiếng Việt
 */
export function normalizePersonName(name: string): {
  original: string;
  cleaned: string;
  noDiacritics: string;
} {
  if (!name || typeof name !== "string") {
    return { original: "", cleaned: "", noDiacritics: "" };
  }

  const cleaned = name
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^a-z0-9\sàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]/g, "");

  const noDiacritics = cleaned
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D");

  return {
    original: name.trim(),
    cleaned,
    noDiacritics,
  };
}

/**
 * Gom danh sách hộ thành các Cụm Địa Bàn (Area Clusters)
 */
export function groupHouseholdsIntoAreaClusters(
  rows: any[],
  colPersonName: string,
  colAreaOrVillage: string,
  colCommune?: string,
  colVillageCode?: string,
  colDbdtName?: string,
  colDbdtCode?: string,
  colExtra?: string
): Map<string, AreaCluster> {
  const clusterMap = new Map<string, AreaCluster>();

  rows.forEach((r, idx) => {
    const rawName = String(r[colPersonName] || "").trim();
    if (!rawName) return;

    const rawArea = String(r[colAreaOrVillage] || "").trim();
    const rawCommune = colCommune ? String(r[colCommune] || "").trim() : "";
    const rawVillageCode = colVillageCode ? String(r[colVillageCode] || "").trim() : "";
    const rawDbdtName = colDbdtName ? String(r[colDbdtName] || "").trim() : "";
    const rawDbdtCode = colDbdtCode ? String(r[colDbdtCode] || "").trim() : "";
    const extraVal = colExtra ? String(r[colExtra] || "").trim() : "";

    // Tạo areaKey duy nhất: kết hợp xã và thôn nếu có
    const areaKey = rawCommune ? `${rawCommune}___${rawArea}` : rawArea || `Khu_vuc_${idx}`;

    const { cleaned, noDiacritics } = normalizePersonName(rawName);

    const item: RawHouseholdItem = {
      raw: r,
      rowIndex: idx,
      fullName: rawName,
      normalizedName: cleaned,
      noDiacriticsName: noDiacritics,
      areaKey,
      communeName: rawCommune,
      communeCode: "",
      villageName: rawArea,
      villageCode: rawVillageCode,
      dbdtName: rawDbdtName,
      dbdtCode: rawDbdtCode,
      extraInfo: extraVal,
    };

    if (!clusterMap.has(areaKey)) {
      clusterMap.set(areaKey, {
        areaKey,
        communeName: rawCommune,
        communeCode: "",
        villageName: rawArea,
        villageCode: rawVillageCode,
        dbdtName: rawDbdtName,
        dbdtCode: rawDbdtCode,
        households: [],
        totalHouseholds: 0,
      });
    }

    const cluster = clusterMap.get(areaKey)!;
    cluster.households.push(item);
    cluster.totalHouseholds = cluster.households.length;
  });

  return clusterMap;
}

/**
 * Thuật toán so khớp ma trận cụm hộ dân
 */
export function executeHouseholdClusterMatching(
  rowsA: any[],
  rowsB: any[],
  config: HouseholdClusterConfig
): {
  results: AreaClusterMatchResult[];
  summary: {
    totalAreasA: number;
    totalAreasB: number;
    matchedAreas: number;
    unmatchedAreas: number;
    matchRate: number;
    totalMatchedHouseholdsAcrossAreas: number;
    thresholdUsed: number;
  };
} {
  const minThreshold = config.minOverlapThreshold || 30;
  const minRatio = config.minOverlapRatio || 0.25;

  // 1. Gom cụm Tệp A và Tệp B
  const clustersA = groupHouseholdsIntoAreaClusters(
    rowsA,
    config.colPersonNameA,
    config.colAreaOrVillageA,
    config.colCommuneA,
    config.colVillageCodeA,
    undefined,
    undefined,
    config.colExtraA
  );

  const clustersB = groupHouseholdsIntoAreaClusters(
    rowsB,
    config.colPersonNameB,
    config.colAreaOrVillageB,
    config.colCommuneB,
    config.colVillageCodeB,
    config.colDbdtNameB,
    config.colDbdtCodeB,
    config.colExtraB
  );

  let clusterListA = Array.from(clustersA.values());
  const clusterListB = Array.from(clustersB.values());

  // Lọc cụm A theo xã nếu người dùng chọn riêng 1 xã
  if (config.selectedCommuneA && config.selectedCommuneA !== "ALL") {
    const selClean = cleanCommuneStr(config.selectedCommuneA);
    clusterListA = clusterListA.filter(cl => cleanCommuneStr(cl.communeName) === selClean);
  }

  // Chuẩn bị bản đồ sáp nhập xã
  const communeRuleMap = new Map<string, string>();
  for (const rule of OFFICIAL_COMMUNE_MERGE_DATA) {
    const k = cleanCommuneStr(rule.communeA);
    const v = cleanCommuneStr(rule.communeB);
    if (k && v) communeRuleMap.set(k, v);
  }

  const results: AreaClusterMatchResult[] = [];
  let totalMatchedHouseholdsAcrossAreas = 0;

  // 2. Chạy đối soát từng cụm địa bàn A với tất cả cụm địa bàn B
  for (const clA of clusterListA) {
    const candidates: AreaClusterCandidate[] = [];

    // Xác định xã mới mục tiêu của clA (nếu có thông tin xã)
    const cleanedCommuneA = cleanCommuneStr(clA.communeName);
    let targetNewCommune = communeRuleMap.get(cleanedCommuneA) || null;
    if (!targetNewCommune && clA.communeName) {
      const found = findOfficialNewCommune(clA.communeName);
      if (found) targetNewCommune = cleanCommuneStr(found);
    }

    // Tạo Index tra cứu tên hộ của clA để so khớp tốc độ cao O(1)
    const nameMapA = new Map<string, RawHouseholdItem[]>();
    for (const hA of clA.households) {
      const k = hA.noDiacriticsName;
      if (!nameMapA.has(k)) nameMapA.set(k, []);
      nameMapA.get(k)!.push(hA);
    }

    for (const clB of clusterListB) {
      // KIỂM TRA PHẠM VI XÃ NẾU BẬT CHẾ ĐỘ NỘI BỘ XÃ:
      const cleanedCommuneB = cleanCommuneStr(clB.communeName);
      if (config.strictCommuneScoped !== false && cleanedCommuneA && cleanedCommuneB) {
        if (targetNewCommune) {
          if (
            cleanedCommuneB !== targetNewCommune &&
            !cleanedCommuneB.includes(targetNewCommune) &&
            !targetNewCommune.includes(cleanedCommuneB)
          ) {
            continue; // Bỏ qua cụm khác xã
          }
        } else {
          if (
            cleanedCommuneA !== cleanedCommuneB &&
            !cleanedCommuneB.includes(cleanedCommuneA) &&
            !cleanedCommuneA.includes(cleanedCommuneB)
          ) {
            continue; // Bỏ qua cụm khác xã
          }
        }
      }

      const matchedPairs: HouseholdMatchedDetail[] = [];
      const usedIndicesA = new Set<number>();

      for (const hB of clB.households) {
        const k = hB.noDiacriticsName;
        const candidatesInA = nameMapA.get(k);

        if (candidatesInA && candidatesInA.length > 0) {
          // Tìm phần tử A chưa ghép
          const targetA = candidatesInA.find(item => !usedIndicesA.has(item.rowIndex));
          if (targetA) {
            usedIndicesA.add(targetA.rowIndex);
            matchedPairs.push({
              nameA: targetA.fullName,
              nameB: hB.fullName,
              rowA: targetA.rowIndex,
              rowB: hB.rowIndex,
              extraA: targetA.extraInfo,
              extraB: hB.extraInfo,
              similarity: 1.0,
            });
          }
        }
      }

      const overlapCount = matchedPairs.length;
      if (overlapCount === 0) continue;

      const overlapRatioA = overlapCount / Math.max(clA.totalHouseholds, 1);
      const overlapRatioB = overlapCount / Math.max(clB.totalHouseholds, 1);

      // Đánh giá mức độ khớp
      let grade: AreaClusterCandidate["matchGrade"] = "LOW_CLUSTER";
      let confidence = 0;
      let reason = "";

      if (overlapCount >= minThreshold) {
        if (overlapCount >= 50 || overlapRatioA >= 0.6) {
          grade = "PERFECT_CLUSTER";
          confidence = 99;
          reason = `Khớp hoàn hảo: Phát hiện ${overlapCount} hộ trùng tên (Vượt xa ngưỡng ${minThreshold} người, đạt ${(overlapRatioA * 100).toFixed(0)}% số hộ thôn cũ)`;
        } else {
          grade = "HIGH_CLUSTER";
          confidence = 95;
          reason = `Khớp cao (Đạt ngưỡng tạm tính): Có ${overlapCount} hộ trùng tên (Ngưỡng yêu cầu: ≥ ${minThreshold} người)`;
        }
      } else if (clA.totalHouseholds < minThreshold && overlapRatioA >= minRatio && overlapCount >= 10) {
        // Trường hợp thôn nhỏ (< 30 hộ) nhưng tỷ lệ trùng rất cao (>= 25-30% và >= 10 người)
        grade = "HIGH_CLUSTER";
        confidence = 88;
        reason = `Thôn nhỏ (${clA.totalHouseholds} hộ): Có ${overlapCount} hộ trùng tên (Đạt ${(overlapRatioA * 100).toFixed(0)}% tổng số hộ của thôn)`;
      } else if (overlapCount >= Math.floor(minThreshold * 0.6) || overlapRatioA >= 0.2) {
        grade = "MEDIUM_CLUSTER";
        confidence = 72;
        reason = `Tiềm năng: Có ${overlapCount} hộ trùng tên (Cần kiểm tra thêm, chưa đạt ngưỡng ${minThreshold} người)`;
      } else {
        grade = "LOW_CLUSTER";
        confidence = 45;
        reason = `Trùng ${overlapCount} hộ (Chưa đủ điều kiện tạm tính)`;
      }

      if (confidence >= 50) {
        candidates.push({
          clusterB: clB,
          overlapCount,
          overlapRatioA,
          overlapRatioB,
          confidence,
          matchGrade: grade,
          matchReason: reason,
          matchedHouseholds: matchedPairs,
        });
      }
    }

    // Sắp xếp các ứng viên B theo số hộ trùng giảm dần
    candidates.sort((a, b) => b.overlapCount - a.overlapCount || b.confidence - a.confidence);

    const best = candidates[0] || null;

    if (best && (best.overlapCount >= minThreshold || best.matchGrade === "HIGH_CLUSTER")) {
      totalMatchedHouseholdsAcrossAreas += best.overlapCount;
      results.push({
        clusterA: clA,
        matchedClusterB: best.clusterB,
        overlapCount: best.overlapCount,
        overlapRatioA: best.overlapRatioA,
        overlapRatioB: best.overlapRatioB,
        confidence: best.confidence,
        matchGrade: best.matchGrade as any,
        matchReason: best.matchReason,
        matchedHouseholds: best.matchedHouseholds,
        candidates,
      });
    } else {
      results.push({
        clusterA: clA,
        matchedClusterB: null,
        overlapCount: best ? best.overlapCount : 0,
        overlapRatioA: best ? best.overlapRatioA : 0,
        overlapRatioB: best ? best.overlapRatioB : 0,
        confidence: 0,
        matchGrade: "UNMATCHED",
        matchReason: best
          ? `Ứng viên cao nhất chỉ trùng ${best.overlapCount} người (Chưa đạt ngưỡng ${minThreshold} người)`
          : "Không tìm thấy cụm hộ dân trùng khớp",
        matchedHouseholds: best ? best.matchedHouseholds : [],
        candidates,
      });
    }
  }

  const matchedAreas = results.filter(r => r.matchedClusterB !== null).length;
  const unmatchedAreas = results.length - matchedAreas;
  const matchRate = results.length > 0 ? Math.round((matchedAreas / results.length) * 100) : 0;

  return {
    results,
    summary: {
      totalAreasA: clusterListA.length,
      totalAreasB: clusterListB.length,
      matchedAreas,
      unmatchedAreas,
      matchRate,
      totalMatchedHouseholdsAcrossAreas,
      thresholdUsed: minThreshold,
    },
  };
}
