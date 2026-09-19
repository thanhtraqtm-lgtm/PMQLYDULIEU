import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  Building2,
  GitMerge,
  Search,
  CheckCircle2,
  AlertTriangle,
  HelpCircle,
  Download,
  ArrowRight,
  Database,
  Sparkles,
  Sliders,
  ChevronDown,
  RefreshCw,
  FileSpreadsheet,
  Layers,
  MapPin,
  Check,
  X,
  UserCheck,
  ClipboardPaste,
  Upload,
  Trash2,
  Plus,
  FileText
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  AdminMergeConfig,
  AdminMergeMatchingResult,
  executeAdminMergeMatching,
  MatchedHouseholdRecord,
  cleanAdminUnitName,
} from "../utils/adminMergeMatcher";
import { VillageAreaCrosswalkResolver } from "./VillageAreaCrosswalkResolver";
import { HouseholdClusterResolver } from "./HouseholdClusterResolver";
import { Users } from "lucide-react";

// Danh mục mẫu 24 cặp Xã cũ ➔ Xã mới theo ảnh chụp công văn sáp nhập
export const SAMPLE_24_COMMUNES: { communeA: string; communeB: string; communeCodeB: string }[] = [
  { communeA: "Xã An Đồng", communeB: "A Sào", communeCodeB: "" },
  { communeA: "Xã An Hiệp", communeB: "A Sào", communeCodeB: "" },
  { communeA: "Xã An Thái", communeB: "A Sào", communeCodeB: "" },
  { communeA: "Xã An Khê", communeB: "A Sào", communeCodeB: "" },
  { communeA: "Xã Tây Giang", communeB: "Ái Quốc", communeCodeB: "" },
  { communeA: "Xã Ái Quốc", communeB: "Ái Quốc", communeCodeB: "" },
  { communeA: "Thị trấn Ân Thi", communeB: "Ân Thi", communeCodeB: "" },
  { communeA: "Xã Quang Vinh", communeB: "Ân Thi", communeCodeB: "" },
  { communeA: "Xã Hoàng Hoa Thám", communeB: "Ân Thi", communeCodeB: "" },
  { communeA: "Xã Đông Cường", communeB: "Bắc Đông Hưng", communeCodeB: "" },
  { communeA: "Xã Đông Xá", communeB: "Bắc Đông Hưng", communeCodeB: "" },
  { communeA: "Xã Đông Phương", communeB: "Bắc Đông Hưng", communeCodeB: "" },
  { communeA: "Xã Hà Giang", communeB: "Bắc Đông Quan", communeCodeB: "" },
  { communeA: "Xã Đông Kinh", communeB: "Bắc Đông Quan", communeCodeB: "" },
  { communeA: "Xã Đông Vinh", communeB: "Bắc Đông Quan", communeCodeB: "" },
  { communeA: "Xã Thái Phúc", communeB: "Bắc Thái Ninh", communeCodeB: "" },
  { communeA: "Xã Dương Hồng Thủy", communeB: "Bắc Thái Ninh", communeCodeB: "" },
  { communeA: "Xã Thụy Quỳnh", communeB: "Bắc Thụy Anh", communeCodeB: "" },
  { communeA: "Xã Thụy Văn", communeB: "Bắc Thụy Anh", communeCodeB: "" },
  { communeA: "Xã Thụy Việt", communeB: "Bắc Thụy Anh", communeCodeB: "" },
  { communeA: "Xã Liên An Đô", communeB: "Bắc Tiên Hưng", communeCodeB: "" },
  { communeA: "Xã Lô Giang", communeB: "Bắc Tiên Hưng", communeCodeB: "" },
  { communeA: "Xã Mê Linh", communeB: "Bắc Tiên Hưng", communeCodeB: "" },
  { communeA: "Xã Phú Lương", communeB: "Bắc Tiên Hưng", communeCodeB: "" },
];

interface AdminMergeResolverProps {
  mainData: any[];
  mainColumns: string[];
  compareData: any[];
  compareColumns: string[];
  compareFileName?: string;
  setMainData?: (data: any[]) => void;
  setColumns?: (cols: string[]) => void;
  onClose?: () => void;
}

export const AdminMergeResolver: React.FC<AdminMergeResolverProps> = ({
  mainData,
  mainColumns,
  compareData,
  compareColumns,
  compareFileName = "Tệp Kỳ 2",
  setMainData,
  setColumns,
  onClose,
}) => {
  // Chế độ xử lý: "village_crosswalk" (Tên thôn + Số hộ), "household_cluster" (≥ 30 người trùng tên), hoặc "household_matching" (Mức hộ gia đình)
  const [resolverMode, setResolverMode] = useState<"village_crosswalk" | "household_cluster" | "household_matching">("village_crosswalk");

  // Cấu hình cột File A (Kỳ NN cũ)
  const [colHouseholdHeadA, setColHouseholdHeadA] = useState<string>("");
  const [colCommuneA, setColCommuneA] = useState<string>("");
  const [colCommuneCodeA, setColCommuneCodeA] = useState<string>("");
  const [colVillageA, setColVillageA] = useState<string>("");
  const [colVillageCodeA, setColVillageCodeA] = useState<string>("");
  const [colIdentifierA, setColIdentifierA] = useState<string>("");

  // Cấu hình cột File B (Kỳ Cá thể mới sau sáp nhập)
  const [colHouseholdHeadB, setColHouseholdHeadB] = useState<string>("");
  const [colCommuneB, setColCommuneB] = useState<string>("");
  const [colCommuneCodeB, setColCommuneCodeB] = useState<string>("");
  const [colVillageB, setColVillageB] = useState<string>("");
  const [colVillageCodeB, setColVillageCodeB] = useState<string>("");
  const [colIdentifierB, setColIdentifierB] = useState<string>("");

  // Bảng quy đổi sáp nhập xã tùy biến do người dùng nhập thêm
  const [customCrosswalk, setCustomCrosswalk] = useState<
    { communeA: string; communeB: string; communeCodeB: string }[]
  >([]);
  const [newRuleA, setNewRuleA] = useState<string>("");
  const [newRuleB, setNewRuleB] = useState<string>("");
  const [newRuleCodeB, setNewRuleCodeB] = useState<string>("");
  const [showPasteModal, setShowPasteModal] = useState<boolean>(false);
  const [pasteRawText, setPasteRawText] = useState<string>("");
  const [crosswalkSearch, setCrosswalkSearch] = useState<string>("");
  const excelInputRef = useRef<HTMLInputElement>(null);

  const filteredCustomCrosswalk = useMemo(() => {
    if (!crosswalkSearch.trim()) return customCrosswalk;
    const q = crosswalkSearch.toLowerCase().trim();
    return customCrosswalk.filter(
      r =>
        r.communeA.toLowerCase().includes(q) ||
        r.communeB.toLowerCase().includes(q) ||
        r.communeCodeB.toLowerCase().includes(q)
    );
  }, [customCrosswalk, crosswalkSearch]);

  // Trạng thái tính toán
  const [matchingResult, setMatchingResult] = useState<AdminMergeMatchingResult | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [activeTab, setActiveTab] = useState<"matched" | "crosswalk" | "review" | "unmatched">("matched");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [filterCommune, setFilterCommune] = useState<string>("ALL");
  const [filterConfidence, setFilterConfidence] = useState<string>("ALL");
  const [appliedMsg, setAppliedMsg] = useState<string | null>(null);

  // Tự động nhận diện cột phù hợp khi nạp dữ liệu
  const autoDetectColumns = () => {
    // 1. Dò File A
    const headA = mainColumns.find(c =>
      /ten.*chu.*ho|chu.*ho|ho.*ten|ten.*ho|ho_va_ten|nguoi_dai_dien|chu_co_so/i.test(c)
    ) || mainColumns.find(c => /ten/i.test(c)) || "";
    setColHouseholdHeadA(headA);

    const commA = mainColumns.find(c => /ten.*xa|xa_phuong|phuong_xa|ten_xa/i.test(c)) ||
      mainColumns.find(c => /xa/i.test(c) && !/ma/i.test(c)) || "";
    setColCommuneA(commA);

    const commCodeA = mainColumns.find(c => /ma.*xa/i.test(c)) || "";
    setColCommuneCodeA(commCodeA);

    const villA = mainColumns.find(c => /ten.*thon|thon|ap|ban|dia_ban|to_dp|tdp|khu_pho/i.test(c) && !/ma/i.test(c)) ||
      mainColumns.find(c => /thon|dia.*ban/i.test(c)) || "";
    setColVillageA(villA);

    const villCodeA = mainColumns.find(c => /ma.*(thon|dia_ban|db)/i.test(c)) || "";
    setColVillageCodeA(villCodeA);

    const idA = mainColumns.find(c => /cccd|cmnd|dinh_danh|sdt|dien_thoai|so_dien_thoai/i.test(c)) || "";
    setColIdentifierA(idA);

    // 2. Dò File B
    const headB = compareColumns.find(c =>
      /ten.*chu.*ho|chu.*ho|ho.*ten|ten.*ho|ten.*co.*so|chu_co_so|ho_va_ten/i.test(c)
    ) || compareColumns.find(c => /ten/i.test(c)) || "";
    setColHouseholdHeadB(headB);

    const commB = compareColumns.find(c => /ten.*xa|xa_phuong|phuong_xa|ten_xa/i.test(c)) ||
      compareColumns.find(c => /xa/i.test(c) && !/ma/i.test(c)) || "";
    setColCommuneB(commB);

    const commCodeB = compareColumns.find(c => /ma.*xa/i.test(c)) || "";
    setColCommuneCodeB(commCodeB);

    const villB = compareColumns.find(c => /ten.*thon|thon|ap|ban|dia_ban|to_dp|tdp|khu_pho/i.test(c) && !/ma/i.test(c)) ||
      compareColumns.find(c => /thon|dia.*ban/i.test(c)) || "";
    setColVillageB(villB);

    const villCodeB = compareColumns.find(c => /ma.*(thon|dia_ban|db)/i.test(c)) || "";
    setColVillageCodeB(villCodeB);

    const idB = compareColumns.find(c => /cccd|cmnd|dinh_danh|sdt|dien_thoai|so_dien_thoai/i.test(c)) || "";
    setColIdentifierB(idB);
  };

  useEffect(() => {
    if (mainColumns.length > 0 && compareColumns.length > 0) {
      autoDetectColumns();
    }
  }, [mainColumns, compareColumns]);

  // Thực hiện thuật toán ghép nối
  const handleRunMatching = () => {
    if (!colHouseholdHeadA || !colHouseholdHeadB) {
      alert("Vui lòng chọn cột Tên chủ hộ ở cả 2 file để hệ thống thực hiện ghép nối.");
      return;
    }

    setIsProcessing(true);
    setAppliedMsg(null);

    // Chuyển customCrosswalk thành dictionary
    const crosswalkDict: Record<string, { communeNameB: string; communeCodeB: string }> = {};
    for (const item of customCrosswalk) {
      const cleanA = cleanAdminUnitName(item.communeA);
      if (cleanA) {
        crosswalkDict[cleanA] = {
          communeNameB: item.communeB,
          communeCodeB: item.communeCodeB,
        };
      }
    }

    const config: AdminMergeConfig = {
      colHouseholdHeadA,
      colCommuneA,
      colCommuneCodeA,
      colVillageA,
      colVillageCodeA,
      colIdentifierA,
      colHouseholdHeadB,
      colCommuneB,
      colCommuneCodeB,
      colVillageB,
      colVillageCodeB,
      colIdentifierB,
      communeCrosswalk: crosswalkDict,
    };

    setTimeout(() => {
      try {
        const res = executeAdminMergeMatching(mainData, compareData, config);
        setMatchingResult(res);
      } catch (err: any) {
        console.error("Error matching:", err);
        alert("Có lỗi xảy ra khi thực hiện ghép nối: " + err.message);
      } finally {
        setIsProcessing(false);
      }
    }, 100);
  };

  // Thêm quy tắc sáp nhập xã thủ công
  const handleAddCustomRule = () => {
    if (!newRuleA.trim() || !newRuleB.trim()) return;
    setCustomCrosswalk([
      ...customCrosswalk,
      {
        communeA: newRuleA.trim(),
        communeB: newRuleB.trim(),
        communeCodeB: newRuleCodeB.trim(),
      },
    ]);
    setNewRuleA("");
    setNewRuleB("");
    setNewRuleCodeB("");
  };

  const handleRemoveCustomRule = (idx: number) => {
    setCustomCrosswalk(customCrosswalk.filter((_, i) => i !== idx));
  };

  // Nạp 24 cặp xã mẫu chụp từ văn bản sáp nhập trong ảnh người dùng gửi
  const handleLoadSampleFromImage = () => {
    const existingMap = new Set(customCrosswalk.map(c => cleanAdminUnitName(c.communeA)));
    const toAdd = SAMPLE_24_COMMUNES.filter(c => !existingMap.has(cleanAdminUnitName(c.communeA)));
    if (toAdd.length === 0) {
      setAppliedMsg("Toàn bộ 24 xã trong bảng mẫu đã có sẵn trong danh sách quy tắc!");
      return;
    }
    setCustomCrosswalk([...customCrosswalk, ...toAdd]);
    setAppliedMsg(`Đã nạp thành công ${toAdd.length} cặp sáp nhập Xã từ bảng mẫu công văn trong ảnh!`);
  };

  // Nạp bảng xã từ text copy-paste trong Excel (Ctrl+V)
  const handleApplyPastedExcel = () => {
    if (!pasteRawText.trim()) return;
    const lines = pasteRawText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    const newRules: { communeA: string; communeB: string; communeCodeB: string }[] = [];

    for (const line of lines) {
      const parts = line.includes("\t") ? line.split("\t") : line.split(/[,;]/);
      if (parts.length >= 2) {
        const cA = parts[0].trim();
        const cB = parts[1].trim();
        const cCodeB = parts[2] ? parts[2].trim() : "";
        // Bỏ qua dòng header nếu có
        if (/xa.*cu|phuong.*cu|ten.*cu/i.test(cA) && /xa.*moi|phuong.*moi|ten.*moi/i.test(cB)) {
          continue;
        }
        if (cA && cB) {
          newRules.push({ communeA: cA, communeB: cB, communeCodeB: cCodeB });
        }
      }
    }

    if (newRules.length === 0) {
      alert("Không tìm thấy dòng nào hợp lệ (cần ít nhất 2 cột: Xã cũ và Xã mới).");
      return;
    }

    const existingMap = new Set(customCrosswalk.map(c => cleanAdminUnitName(c.communeA)));
    const filtered = newRules.filter(r => !existingMap.has(cleanAdminUnitName(r.communeA)));
    setCustomCrosswalk([...customCrosswalk, ...filtered]);
    setShowPasteModal(false);
    setPasteRawText("");
    setAppliedMsg(`Đã nạp thêm ${filtered.length} quy tắc sáp nhập xã từ bảng bạn vừa dán!`);
  };

  // Nạp từ tệp Excel / CSV bảng sáp nhập
  const handleImportExcelCrosswalkFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const firstSheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[firstSheetName];
      const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      if (rawRows.length < 2) {
        alert("Tệp Excel bảng sáp nhập không có đủ dữ liệu.");
        return;
      }

      let headerRowIdx = 0;
      let colAIdx = 0;
      let colBIdx = 1;
      let colCodeBIdx = -1;

      for (let r = 0; r < Math.min(5, rawRows.length); r++) {
        const row = rawRows[r] || [];
        const foundA = row.findIndex((val: any) => /xa.*cu|phuong.*cu|ten.*cu|cu$/i.test(String(val || "")));
        const foundB = row.findIndex((val: any) => /xa.*moi|phuong.*moi|ten.*moi|moi$/i.test(String(val || "")));
        if (foundA !== -1 && foundB !== -1) {
          headerRowIdx = r;
          colAIdx = foundA;
          colBIdx = foundB;
          colCodeBIdx = row.findIndex((val: any) => /ma.*xa.*moi|ma.*moi/i.test(String(val || "")));
          break;
        }
      }

      const newRules: { communeA: string; communeB: string; communeCodeB: string }[] = [];
      for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r] || [];
        const cA = String(row[colAIdx] || "").trim();
        const cB = String(row[colBIdx] || "").trim();
        const cCodeB = colCodeBIdx !== -1 ? String(row[colCodeBIdx] || "").trim() : "";
        if (cA && cB && !/xa.*cu/i.test(cA)) {
          newRules.push({ communeA: cA, communeB: cB, communeCodeB: cCodeB });
        }
      }

      if (newRules.length === 0) {
        alert("Không tìm thấy dữ liệu Xã cũ và Xã mới hợp lệ trong tệp Excel.");
        return;
      }

      const existingMap = new Set(customCrosswalk.map(c => cleanAdminUnitName(c.communeA)));
      const filtered = newRules.filter(r => !existingMap.has(cleanAdminUnitName(r.communeA)));
      setCustomCrosswalk([...customCrosswalk, ...filtered]);
      setAppliedMsg(`Đã nạp thành công ${filtered.length} quy tắc sáp nhập xã từ tệp "${file.name}"!`);
    } catch (err: any) {
      alert("Lỗi khi đọc file Excel: " + (err?.message || "Không xác định"));
    } finally {
      e.target.value = "";
    }
  };

  const handleClearAllRules = () => {
    if (customCrosswalk.length === 0) return;
    if (confirm("Bạn có chắc chắn muốn xóa toàn bộ quy tắc sáp nhập xã đã nạp?")) {
      setCustomCrosswalk([]);
    }
  };

  // Áp dụng gán mã xã/thôn mới ngược lại vào mainData (File A)
  const handleApplyCodesToMainData = () => {
    if (!matchingResult || !setMainData) return;

    let updatedFullCount = 0;
    let updatedCommuneOnlyCount = 0;

    const updatedMainData = mainData.map((row, idx) => {
      const match = matchingResult.records[idx];
      const newRow = { ...row };

      if (match && match.status !== "unmatched") {
        newRow["Ma_Xa_Moi"] = match.communeCodeB || "";
        newRow["Ten_Xa_Moi"] = match.communeB || "";
        newRow["Ma_Thon_Moi"] = match.villageCodeB || "";
        newRow["Ten_Thon_Moi"] = match.villageB || "";
        newRow["Trang_Thai_Ghep_Sap_Nhap"] =
          match.status === "exact_id"
            ? "Khớp CCCD/SĐT (100%)"
            : match.status === "high_match"
            ? `Khớp cao (${match.confidence}%)`
            : `Khớp trung bình (${match.confidence}%)`;
        newRow["Can_Cu_Ghep"] = match.matchReason;
        updatedFullCount++;
      } else {
        // Nếu chưa khớp được hộ nhưng đã định vị được xã mới (qua bảng sáp nhập)
        const hasMappedCommune = Boolean(match?.communeB);
        newRow["Ma_Xa_Moi"] = match?.communeCodeB || "";
        newRow["Ten_Xa_Moi"] = match?.communeB || "";
        newRow["Ma_Thon_Moi"] = "";
        newRow["Ten_Thon_Moi"] = "";
        newRow["Trang_Thai_Ghep_Sap_Nhap"] = hasMappedCommune
          ? "Đã định vị Xã mới (Chưa khớp Thôn/Hộ)"
          : "Chưa tìm thấy trên địa bàn mới";
        newRow["Can_Cu_Ghep"] = match?.matchReason || "";
        if (hasMappedCommune) {
          updatedCommuneOnlyCount++;
        }
      }

      return newRow;
    });

    // Cập nhật danh sách cột nếu có hàm setColumns
    if (setColumns) {
      const addedCols = ["Ma_Xa_Moi", "Ten_Xa_Moi", "Ma_Thon_Moi", "Ten_Thon_Moi", "Trang_Thai_Ghep_Sap_Nhap", "Can_Cu_Ghep"];
      const newCols = [...mainColumns];
      addedCols.forEach(c => {
        if (!newCols.includes(c)) newCols.push(c);
      });
      setColumns(newCols);
    }

    setMainData(updatedMainData);
    let msg = `Đã gán Mã Xã Mới & Mã Thôn Mới cho ${updatedFullCount.toLocaleString("vi-VN")} hộ!`;
    if (updatedCommuneOnlyCount > 0) {
      msg += ` Đồng thời gán Tên/Mã Xã Mới cho ${updatedCommuneOnlyCount.toLocaleString("vi-VN")} hộ khác dựa theo bảng sáp nhập.`;
    }
    setAppliedMsg(msg);
  };

  // Chọn thủ công một ứng viên trong danh sách nghi ngờ
  const handleSelectCandidateManually = (recordIndexA: number, candRowB: any) => {
    if (!matchingResult) return;

    const newRecords = [...matchingResult.records];
    const rec = newRecords[recordIndexA];
    if (!rec) return;

    const commB = colCommuneB ? String(candRowB[colCommuneB] || "").trim() : "";
    const commCodeB = colCommuneCodeB ? String(candRowB[colCommuneCodeB] || "").trim() : "";
    const villB = colVillageB ? String(candRowB[colVillageB] || "").trim() : "";
    const villCodeB = colVillageCodeB ? String(candRowB[colVillageCodeB] || "").trim() : "";
    const headB = String(candRowB[colHouseholdHeadB] || "").trim();

    newRecords[recordIndexA] = {
      ...rec,
      rowB: candRowB,
      householdHeadB: headB,
      communeB: commB,
      communeCodeB: commCodeB,
      villageB: villB,
      villageCodeB: villCodeB,
      status: "high_match",
      matchReason: `Xác nhận thủ công bởi cán bộ: ${headB} (${commB} - ${villB})`,
      confidence: 95,
    };

    // Cập nhật lại stats
    const matchedCount = newRecords.filter(r => r.status !== "unmatched").length;
    setMatchingResult({
      ...matchingResult,
      records: newRecords,
      stats: {
        ...matchingResult.stats,
        matchedCount,
        unmatchedCount: matchingResult.stats.totalA - matchedCount,
        matchRate: (matchedCount / matchingResult.stats.totalA) * 100,
      },
    });
  };

  // Xuất file Excel chuyên nghiệp gồm 4 Sheet
  const handleExportFullExcel = () => {
    if (!matchingResult) return;

    const wb = XLSX.utils.book_new();

    // Sheet 1: Danh sách hộ đã ghép
    const matchedDataRows = matchingResult.records
      .filter(r => r.status !== "unmatched")
      .map((r, i) => ({
        STT: i + 1,
        "Tên chủ hộ (Kỳ 1 NN cũ)": r.householdHeadA,
        "Xã cũ (Kỳ 1)": r.communeA,
        "Mã xã cũ": r.communeCodeA,
        "Thôn cũ (Kỳ 1)": r.villageA,
        "Mã thôn cũ": r.villageCodeA,
        "CCCD/SĐT (Kỳ 1)": r.identifierA,
        "--- ĐỊA BÀN MỚI SAU SÁP NHẬP ---": "===>",
        "Tên chủ/cơ sở (Kỳ 2 mới)": r.householdHeadB,
        "Xã mới (Kỳ 2)": r.communeB,
        "MÃ XÃ MỚI": r.communeCodeB,
        "Thôn/Địa bàn mới": r.villageB,
        "MÃ THÔN MỚI": r.villageCodeB,
        "Độ tin cậy (%)": r.confidence,
        "Trạng thái":
          r.status === "exact_id"
            ? "Khớp CCCD/SĐT"
            : r.status === "high_match"
            ? "Khớp cao"
            : "Khớp trung bình",
        "Căn cứ ghép nối": r.matchReason,
      }));

    const ws1 = XLSX.utils.json_to_sheet(matchedDataRows);
    XLSX.utils.book_append_sheet(wb, ws1, "Ho_Da_Ghep_Dia_Ban_Moi");

    // Sheet 2: Danh sách chưa tìm thấy
    const unmatchedDataRows = matchingResult.records
      .filter(r => r.status === "unmatched")
      .map((r, i) => ({
        STT: i + 1,
        "Tên chủ hộ (Kỳ 1 NN cũ)": r.householdHeadA,
        "Xã cũ": r.communeA,
        "Mã xã cũ": r.communeCodeA,
        "Thôn cũ": r.villageA,
        "Mã thôn cũ": r.villageCodeA,
        "CCCD/SĐT": r.identifierA,
        "Ghi chú kiểm tra": r.matchReason,
        "Gợi ý gần nhất": r.otherCandidates?.[0]
          ? `${r.otherCandidates[0].householdHeadB} (${r.otherCandidates[0].communeB} - ${r.otherCandidates[0].villageB})`
          : "Không có",
      }));

    const ws2 = XLSX.utils.json_to_sheet(unmatchedDataRows);
    XLSX.utils.book_append_sheet(wb, ws2, "Ho_Chua_Tim_Thay");

    // Sheet 3: Bảng ánh xạ Xã sáp nhập
    const communeCrossRows = matchingResult.communeCrosswalk.map((c, i) => ({
      STT: i + 1,
      "Tên Xã Cũ (Trước sáp nhập)": c.communeNameA,
      "Mã Xã Cũ": c.communeCodeA,
      "SÁP NHẬP VÀO": "===>",
      "Tên Xã Mới (Sau sáp nhập)": c.communeNameB,
      "Mã Xã Mới": c.communeCodeB,
      "Số hộ dân khớp được": c.matchedCount,
      "Độ tin cậy quy đổi (%)": c.confidence,
    }));

    const ws3 = XLSX.utils.json_to_sheet(communeCrossRows);
    XLSX.utils.book_append_sheet(wb, ws3, "Anh_Xa_Xa_Sap_Nhap");

    // Sheet 4: Bảng ánh xạ Thôn sáp nhập
    const villageCrossRows = matchingResult.villageCrosswalk.map((v, i) => ({
      STT: i + 1,
      "Xã Cũ": v.communeNameA,
      "Thôn Cũ (Kỳ 1)": v.villageNameA,
      "Mã Thôn Cũ": v.villageCodeA,
      "SÁP NHẬP VÀO": "===>",
      "Xã Mới": v.communeNameB,
      "Thôn Mới (Kỳ 2)": v.villageNameB,
      "Mã Thôn Mới": v.villageCodeB,
      "Số hộ dân khớp được": v.matchedCount,
    }));

    const ws4 = XLSX.utils.json_to_sheet(villageCrossRows);
    XLSX.utils.book_append_sheet(wb, ws4, "Anh_Xa_Thon_Sap_Nhap");

    XLSX.writeFile(wb, `Doi_Soat_Sap_Nhap_Ho_NN_va_Dia_Ban_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Danh sách các Xã cũ để làm bộ lọc
  const uniqueCommunesA = useMemo(() => {
    if (!matchingResult) return [];
    const set = new Set<string>();
    matchingResult.records.forEach(r => {
      if (r.communeA) set.add(r.communeA);
    });
    return Array.from(set).sort();
  }, [matchingResult]);

  // Danh sách lọc hiển thị theo tab
  const displayedRecords = useMemo(() => {
    if (!matchingResult) return [];

    let list = matchingResult.records;

    // Lọc theo Tab
    if (activeTab === "matched") {
      list = list.filter(r => r.status !== "unmatched");
    } else if (activeTab === "review") {
      list = list.filter(
        r => r.status === "medium_match" || (r.status === "unmatched" && (r.otherCandidates?.length || 0) > 0)
      );
    } else if (activeTab === "unmatched") {
      list = list.filter(r => r.status === "unmatched");
    }

    // Lọc theo Xã
    if (filterCommune !== "ALL") {
      list = list.filter(r => r.communeA === filterCommune);
    }

    // Lọc theo Độ tin cậy
    if (filterConfidence === "EXACT") {
      list = list.filter(r => r.status === "exact_id");
    } else if (filterConfidence === "HIGH") {
      list = list.filter(r => r.status === "high_match");
    } else if (filterConfidence === "MEDIUM") {
      list = list.filter(r => r.status === "medium_match");
    }

    // Lọc theo tìm kiếm từ khóa
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      list = list.filter(
        r =>
          r.householdHeadA.toLowerCase().includes(q) ||
          r.communeA.toLowerCase().includes(q) ||
          r.villageA.toLowerCase().includes(q) ||
          (r.householdHeadB && r.householdHeadB.toLowerCase().includes(q)) ||
          (r.communeB && r.communeB.toLowerCase().includes(q)) ||
          (r.villageB && r.villageB.toLowerCase().includes(q)) ||
          r.identifierA.toLowerCase().includes(q)
      );
    }

    return list;
  }, [matchingResult, activeTab, filterCommune, filterConfidence, searchQuery]);

  if (resolverMode === "village_crosswalk") {
    return (
      <div className="space-y-3 font-sans">
        {/* THANH CHUYỂN ĐỔI CHẾ ĐỘ ĐỐI SOÁT */}
        <div className="bg-slate-900 text-white p-2 flex flex-wrap items-center justify-between gap-2 border border-slate-700 shadow-sm">
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-xs font-bold uppercase text-amber-300">Chế độ:</span>
            <div className="inline-flex border border-slate-700 overflow-hidden shadow-xs">
              <button
                type="button"
                onClick={() => setResolverMode("village_crosswalk")}
                className="px-2.5 py-1 text-xs font-bold bg-[#286e42] text-white border-0 cursor-pointer flex items-center gap-1 shadow-inner"
              >
                <MapPin className="w-3.5 h-3.5 text-amber-300" />
                <span>1. Đối soát địa bàn &amp; thôn</span>
              </button>
              <button
                type="button"
                onClick={() => setResolverMode("household_cluster")}
                className="px-2.5 py-1 text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-300 border-0 cursor-pointer flex items-center gap-1 transition-colors"
              >
                <Users className="w-3.5 h-3.5 text-amber-400" />
                <span>2. So cụm hộ dân</span>
              </button>
              <button
                type="button"
                onClick={() => setResolverMode("household_matching")}
                className="px-2.5 py-1 text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-300 border-0 cursor-pointer flex items-center gap-1 transition-colors"
              >
                <UserCheck className="w-3.5 h-3.5 text-sky-400" />
                <span>3. Đối soát chủ hộ</span>
              </button>
            </div>
          </div>
          {onClose && (
            <button
              type="button"
              onClick={onClose}
              className="bg-white/10 hover:bg-white/20 text-white text-xs px-2.5 py-1 border border-white/20 cursor-pointer"
            >
              ✕ Đóng
            </button>
          )}
        </div>

        <VillageAreaCrosswalkResolver
          initialDataA={mainData}
          initialColumnsA={mainColumns}
          initialDataB={compareData}
          initialColumnsB={compareColumns}
          onApplyMergedToMain={(merged, cols) => {
            if (setMainData) setMainData(merged);
            if (setColumns) setColumns(cols);
          }}
          onClose={onClose}
        />
      </div>
    );
  }

  if (resolverMode === "household_cluster") {
    return (
      <div className="space-y-3 font-sans">
        {/* THANH CHUYỂN ĐỔI CHẾ ĐỘ ĐỐI SOÁT */}
        <div className="bg-slate-900 text-white p-2 flex flex-wrap items-center justify-between gap-2 border border-slate-700 shadow-sm">
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-xs font-bold uppercase text-amber-300">Chế độ:</span>
            <div className="inline-flex border border-slate-700 overflow-hidden shadow-xs">
              <button
                type="button"
                onClick={() => setResolverMode("village_crosswalk")}
                className="px-2.5 py-1 text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-300 border-0 cursor-pointer flex items-center gap-1 transition-colors"
              >
                <MapPin className="w-3.5 h-3.5 text-amber-300" />
                <span>1. Đối soát địa bàn &amp; thôn</span>
              </button>
              <button
                type="button"
                onClick={() => setResolverMode("household_cluster")}
                className="px-2.5 py-1 text-xs font-bold bg-[#1e5430] text-white border-0 cursor-pointer flex items-center gap-1 shadow-inner"
              >
                <Users className="w-3.5 h-3.5 text-amber-300" />
                <span>2. So cụm hộ dân</span>
              </button>
              <button
                type="button"
                onClick={() => setResolverMode("household_matching")}
                className="px-2.5 py-1 text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-300 border-0 cursor-pointer flex items-center gap-1 transition-colors"
              >
                <UserCheck className="w-3.5 h-3.5 text-sky-400" />
                <span>3. Đối soát chủ hộ</span>
              </button>
            </div>
          </div>
          {onClose && (
            <button
              type="button"
              onClick={onClose}
              className="bg-white/10 hover:bg-white/20 text-white text-xs px-2.5 py-1 border border-white/20 cursor-pointer"
            >
              ✕ Đóng
            </button>
          )}
        </div>

        <HouseholdClusterResolver
          initialDataA={mainData}
          initialColumnsA={mainColumns}
          initialDataB={compareData}
          initialColumnsB={compareColumns}
          onApplyMergedToMain={(merged, cols) => {
            if (setMainData) setMainData(merged);
            if (setColumns) setColumns(cols);
          }}
          onClose={onClose}
        />
      </div>
    );
  }

  return (
    <div className="bg-slate-50 border-2 border-indigo-400 rounded-none shadow-md overflow-hidden space-y-3 font-sans">
      {/* THANH CHUYỂN ĐỔI CHẾ ĐỘ ĐỐI SOÁT */}
      <div className="bg-slate-900 text-white p-2 flex flex-wrap items-center justify-between gap-2 border-b border-slate-700 shadow-sm">
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-xs font-bold uppercase text-amber-300">Chế độ:</span>
          <div className="inline-flex border border-slate-700 overflow-hidden shadow-xs">
            <button
              type="button"
              onClick={() => setResolverMode("village_crosswalk")}
              className="px-2.5 py-1 text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-300 border-0 cursor-pointer flex items-center gap-1 transition-colors"
            >
              <MapPin className="w-3.5 h-3.5 text-amber-300" />
              <span>1. Đối soát địa bàn &amp; thôn</span>
            </button>
            <button
              type="button"
              onClick={() => setResolverMode("household_cluster")}
              className="px-2.5 py-1 text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-300 border-0 cursor-pointer flex items-center gap-1 transition-colors"
            >
              <Users className="w-3.5 h-3.5 text-amber-400" />
              <span>2. So cụm hộ dân</span>
            </button>
            <button
              type="button"
              onClick={() => setResolverMode("household_matching")}
              className="px-2.5 py-1 text-xs font-bold bg-indigo-700 text-white border-0 cursor-pointer flex items-center gap-1 shadow-inner"
            >
              <UserCheck className="w-3.5 h-3.5 text-sky-400" />
              <span>3. Đối soát chủ hộ</span>
            </button>
          </div>
        </div>
        {onClose && (
          <button
            type="button"
            onClick={onClose}
            className="bg-white/10 hover:bg-white/20 text-white text-xs px-2.5 py-1 border border-white/20 cursor-pointer"
          >
            ✕ Đóng
          </button>
        )}
      </div>

      {/* 1. THANH TIÊU ĐỀ NỔI BẬT */}
      <div className="bg-gradient-to-r from-indigo-900 via-indigo-800 to-sky-900 px-4 py-3 text-white flex flex-col md:flex-row md:items-center justify-between gap-3 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 bg-amber-500 text-slate-950 flex items-center justify-center font-bold shadow-sm">
            <Building2 className="w-5 h-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h3 className="text-sm font-bold tracking-wide uppercase font-mono">
                Xử Lý Sáp Nhập Xã / Thôn: Hộ Điều Tra NN ↔ Địa Bàn Kỳ Cá Thể
              </h3>
              <span className="bg-amber-400 text-slate-950 text-[10px] font-extrabold px-1.5 py-0.5 uppercase tracking-wider">
                Thuật toán Vượt Rào Cản Mất Mã
              </span>
            </div>
            <p className="text-xs text-indigo-200 mt-0.5">
              Tự động dò tìm địa bàn mới cho từng hộ dân khi các xã, thôn bị gộp/đổi mã; trích xuất Bảng ánh xạ xã thôn và gán mã mới vào danh sách cũ.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {matchingResult && (
            <>
              <button
                type="button"
                onClick={handleApplyCodesToMainData}
                className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1.5 shadow-xs border-0 cursor-pointer"
                title="Tự động thêm các cột Ma_Xa_Moi, Ten_Xa_Moi, Ma_Thon_Moi, Ten_Thon_Moi vào bảng làm việc"
              >
                <Database className="w-4 h-4" />
                <span>Gán Mã Mới Vào Dữ Liệu Kỳ 1</span>
              </button>

              <button
                type="button"
                onClick={handleExportFullExcel}
                className="bg-amber-500 hover:bg-amber-600 text-slate-950 font-bold text-xs px-3 py-1.5 flex items-center gap-1.5 shadow-xs border-0 cursor-pointer"
                title="Tải báo cáo Excel gồm 4 sheet: Hộ đã ghép, Hộ chưa tìm thấy, Ánh xạ Xã, Ánh xạ Thôn"
              >
                <Download className="w-4 h-4" />
                <span>Xuất Excel Đầy Đủ (4 Sheet)</span>
              </button>
            </>
          )}

          {onClose && (
            <button
              type="button"
              onClick={onClose}
              className="bg-white/10 hover:bg-white/20 text-white text-xs font-bold px-2 py-1.5 border border-white/20 cursor-pointer"
            >
              ✕ Đóng
            </button>
          )}
        </div>
      </div>

      {/* THÔNG BÁO GÁN MÃ THÀNH CÔNG */}
      {appliedMsg && (
        <div className="mx-4 p-3 bg-emerald-100 border-2 border-emerald-400 text-emerald-950 text-xs font-bold flex items-center justify-between gap-2 shadow-xs">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-5 h-5 text-emerald-700 shrink-0" />
            <span>{appliedMsg}</span>
          </div>
          <button
            type="button"
            onClick={() => setAppliedMsg(null)}
            className="text-emerald-800 hover:text-black font-bold border-0 bg-transparent cursor-pointer"
          >
            ✕
          </button>
        </div>
      )}

      {/* 2. KHU VỰC CẤU HÌNH CỘT 2 TỆP (INPUT MAPPING) */}
      <div className="mx-4 bg-white border border-indigo-200 p-3.5 space-y-3 shadow-2xs">
        <div className="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-indigo-100">
          <div className="flex items-center gap-2">
            <Sliders className="w-4 h-4 text-indigo-700" />
            <span className="text-xs font-bold text-slate-900 uppercase">
              1. Cấu hình các cột tương ứng giữa 2 kỳ điều tra
            </span>
          </div>
          <button
            type="button"
            onClick={autoDetectColumns}
            className="text-[11px] font-bold text-indigo-700 hover:text-indigo-900 bg-indigo-50 hover:bg-indigo-100 px-2 py-1 border border-indigo-300 flex items-center gap-1 cursor-pointer"
          >
            <Sparkles className="w-3.5 h-3.5 text-amber-600" />
            <span>⚡ Tự động nhận diện lại cột</span>
          </button>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {/* Cột Trái: File A - Điều tra NN (cũ) */}
          <div className="bg-sky-50/50 p-3 border border-sky-200 space-y-2.5">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-sky-950 flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-sky-600"></span>
                KỲ 1: DANH SÁCH HỘ (ĐIỀU TRA NN CŨ)
              </span>
              <span className="text-[11px] font-mono text-sky-800 bg-sky-100 px-1.5 py-0.5 border border-sky-300">
                {mainData.length.toLocaleString("vi-VN")} hộ
              </span>
            </div>

            <div className="space-y-2">
              <div>
                <label className="block text-[11px] font-bold text-slate-800 mb-0.5">
                  ⭐ Cột Tên chủ hộ / Người đại diện (Bắt buộc):
                </label>
                <select
                  value={colHouseholdHeadA}
                  onChange={e => setColHouseholdHeadA(e.target.value)}
                  className="w-full text-xs border border-sky-300 bg-white px-2 py-1 font-bold text-slate-900 focus:ring-1 focus:ring-indigo-500"
                >
                  <option value="">-- Chọn cột Tên chủ hộ Kỳ 1 --</option>
                  {mainColumns.map(c => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                    Tên Xã cũ (Kỳ 1):
                  </label>
                  <select
                    value={colCommuneA}
                    onChange={e => setColCommuneA(e.target.value)}
                    className="w-full text-xs border border-slate-300 bg-white px-1.5 py-1"
                  >
                    <option value="">-- Tên Xã cũ --</option>
                    {mainColumns.map(c => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                    Mã Xã cũ (nếu có):
                  </label>
                  <select
                    value={colCommuneCodeA}
                    onChange={e => setColCommuneCodeA(e.target.value)}
                    className="w-full text-xs border border-slate-300 bg-white px-1.5 py-1"
                  >
                    <option value="">-- Không có mã --</option>
                    {mainColumns.map(c => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                    Tên Thôn / Bản / TDP cũ:
                  </label>
                  <select
                    value={colVillageA}
                    onChange={e => setColVillageA(e.target.value)}
                    className="w-full text-xs border border-slate-300 bg-white px-1.5 py-1"
                  >
                    <option value="">-- Tên Thôn cũ --</option>
                    {mainColumns.map(c => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                    Mã Thôn cũ (nếu có):
                  </label>
                  <select
                    value={colVillageCodeA}
                    onChange={e => setColVillageCodeA(e.target.value)}
                    className="w-full text-xs border border-slate-300 bg-white px-1.5 py-1"
                  >
                    <option value="">-- Không có mã --</option>
                    {mainColumns.map(c => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                  CCCD / CMND / SĐT (nếu có):
                </label>
                <select
                  value={colIdentifierA}
                  onChange={e => setColIdentifierA(e.target.value)}
                  className="w-full text-xs border border-slate-300 bg-white px-1.5 py-1 text-slate-700"
                >
                  <option value="">-- Không có (Khớp theo tên &amp; thôn) --</option>
                  {mainColumns.map(c => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* Cột Phải: File B - Địa bàn / Cá thể (mới sau sáp nhập) */}
          <div className="bg-emerald-50/50 p-3 border border-emerald-200 space-y-2.5">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-emerald-950 flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-emerald-600"></span>
                KỲ 2: ĐỊA BÀN / CÁ THỂ (ĐÃ SÁP NHẬP NĂM NAY)
              </span>
              <span className="text-[11px] font-mono text-emerald-800 bg-emerald-100 px-1.5 py-0.5 border border-emerald-300">
                {compareData.length.toLocaleString("vi-VN")} dòng
              </span>
            </div>

            <div className="space-y-2">
              <div>
                <label className="block text-[11px] font-bold text-slate-800 mb-0.5">
                  ⭐ Cột Tên chủ cơ sở / Tên hộ / Tên đơn vị (Bắt buộc):
                </label>
                <select
                  value={colHouseholdHeadB}
                  onChange={e => setColHouseholdHeadB(e.target.value)}
                  className="w-full text-xs border border-emerald-300 bg-white px-2 py-1 font-bold text-slate-900 focus:ring-1 focus:ring-emerald-500"
                >
                  <option value="">-- Chọn cột Tên chủ/hộ Kỳ 2 --</option>
                  {compareColumns.map(c => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                    Tên Xã MỚI (Kỳ 2):
                  </label>
                  <select
                    value={colCommuneB}
                    onChange={e => setColCommuneB(e.target.value)}
                    className="w-full text-xs border border-slate-300 bg-white px-1.5 py-1"
                  >
                    <option value="">-- Tên Xã mới --</option>
                    {compareColumns.map(c => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5 text-emerald-800">
                    Mã Xã MỚI (để lấy về):
                  </label>
                  <select
                    value={colCommuneCodeB}
                    onChange={e => setColCommuneCodeB(e.target.value)}
                    className="w-full text-xs border border-emerald-400 bg-emerald-50 font-bold px-1.5 py-1 text-emerald-950"
                  >
                    <option value="">-- Chọn cột Mã xã mới --</option>
                    {compareColumns.map(c => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                    Tên Thôn / Địa bàn MỚI:
                  </label>
                  <select
                    value={colVillageB}
                    onChange={e => setColVillageB(e.target.value)}
                    className="w-full text-xs border border-slate-300 bg-white px-1.5 py-1"
                  >
                    <option value="">-- Tên Thôn mới --</option>
                    {compareColumns.map(c => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-700 mb-0.5 text-emerald-800">
                    Mã Thôn MỚI (để lấy về):
                  </label>
                  <select
                    value={colVillageCodeB}
                    onChange={e => setColVillageCodeB(e.target.value)}
                    className="w-full text-xs border border-emerald-400 bg-emerald-50 font-bold px-1.5 py-1 text-emerald-950"
                  >
                    <option value="">-- Chọn cột Mã thôn mới --</option>
                    {compareColumns.map(c => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-bold text-slate-700 mb-0.5">
                  CCCD / CMND / SĐT (Kỳ 2):
                </label>
                <select
                  value={colIdentifierB}
                  onChange={e => setColIdentifierB(e.target.value)}
                  className="w-full text-xs border border-slate-300 bg-white px-1.5 py-1 text-slate-700"
                >
                  <option value="">-- Không có --</option>
                  {compareColumns.map(c => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        </div>

        {/* 3. BẢNG QUY TẮC SÁP NHẬP XÃ BỔ SUNG (NẾU ĐÃ BIẾT TRƯỚC NGHỊ QUYẾT SÁP NHẬP) */}
        <div className="bg-amber-50/80 border border-amber-300 p-2.5 space-y-2.5">
          {/* Header với các nút nạp nhanh */}
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div>
              <span className="text-[11px] font-bold text-amber-950 flex items-center gap-1.5">
                <GitMerge className="w-4 h-4 text-amber-700" />
                Bảng quy tắc sáp nhập Xã đã biết (Khuyên dùng: Đạt độ chính xác 100%):
              </span>
              <p className="text-[10px] text-amber-800 m-0">
                {customCrosswalk.length > 0
                  ? `Đang áp dụng ${customCrosswalk.length} quy tắc sáp nhập xã (Hệ thống sẽ chỉ dò tìm hộ cũ trong đúng xã mới được quy định).`
                  : "Chưa nạp quy tắc. Hệ thống sẽ tự động học và suy luận quy luật sáp nhập qua các hộ khớp CCCD/SĐT."}
              </p>
            </div>

            {/* Các nút thao tác nạp nhanh */}
            <div className="flex flex-wrap items-center gap-1.5">
              <button
                type="button"
                onClick={() => setShowPasteModal(true)}
                className="bg-sky-700 hover:bg-sky-800 text-white font-bold text-[11px] px-2.5 py-1 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                title="Sao chép cột Xã cũ và Xã mới trong Excel rồi dán (Ctrl+V) vào đây"
              >
                <ClipboardPaste className="w-3.5 h-3.5" />
                📋 Dán nhanh từ Excel
              </button>

              <input
                type="file"
                ref={excelInputRef}
                accept=".xlsx,.xls,.csv"
                onChange={handleImportExcelCrosswalkFile}
                className="hidden"
              />
              <button
                type="button"
                onClick={() => excelInputRef.current?.click()}
                className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-[11px] px-2.5 py-1 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                title="Tải lên tệp Excel chứa bảng quy đổi sáp nhập xã"
              >
                <Upload className="w-3.5 h-3.5" />
                📥 Nạp tệp Excel bảng sáp nhập
              </button>

              <button
                type="button"
                onClick={handleLoadSampleFromImage}
                className="bg-amber-600 hover:bg-amber-700 text-white font-bold text-[11px] px-2.5 py-1 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                title="Nạp nhanh 24 cặp Xã theo ảnh chụp công văn sáp nhập bạn vừa gửi"
              >
                <Sparkles className="w-3.5 h-3.5" />
                ⚡ Nạp 24 Xã từ ảnh của bạn
              </button>

              {customCrosswalk.length > 0 && (
                <button
                  type="button"
                  onClick={handleClearAllRules}
                  className="bg-slate-200 hover:bg-red-100 text-slate-700 hover:text-red-700 font-bold text-[11px] px-2 py-1 flex items-center gap-1 cursor-pointer border border-slate-300"
                  title="Xóa toàn bộ các quy tắc hiện tại"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  Xóa tất cả ({customCrosswalk.length})
                </button>
              )}
            </div>
          </div>

          {/* Dòng thêm thủ công 1 cặp xã */}
          <div className="flex flex-wrap items-center gap-2 bg-amber-100/50 p-2 border border-amber-200">
            <span className="text-[11px] font-bold text-amber-900">Thêm lẻ 1 xã:</span>
            <input
              type="text"
              placeholder="Tên Xã cũ (VD: Xã An Đồng)"
              value={newRuleA}
              onChange={e => setNewRuleA(e.target.value)}
              className="text-xs border border-amber-300 bg-white px-2 py-1 w-44"
            />
            <span className="text-xs font-bold text-amber-700">➔ Sáp nhập vào:</span>
            <input
              type="text"
              placeholder="Tên Xã mới (VD: A Sào)"
              value={newRuleB}
              onChange={e => setNewRuleB(e.target.value)}
              className="text-xs border border-amber-300 bg-white px-2 py-1 w-44"
            />
            <input
              type="text"
              placeholder="Mã Xã mới (nếu có)"
              value={newRuleCodeB}
              onChange={e => setNewRuleCodeB(e.target.value)}
              className="text-xs border border-amber-300 bg-white px-2 py-1 w-28"
            />
            <button
              type="button"
              onClick={handleAddCustomRule}
              className="bg-amber-700 hover:bg-amber-800 text-white font-bold text-xs px-2.5 py-1 cursor-pointer border-0 flex items-center gap-1"
            >
              <Plus className="w-3.5 h-3.5" /> Thêm quy tắc
            </button>
          </div>

          {/* Danh sách các quy tắc đã nạp kèm ô tìm kiếm */}
          {customCrosswalk.length > 0 && (
            <div className="space-y-1.5 pt-1">
              <div className="flex items-center justify-between">
                <span className="text-[10.5px] font-bold text-amber-900">
                  Danh mục đã nạp ({filteredCustomCrosswalk.length} / {customCrosswalk.length} quy tắc):
                </span>
                {customCrosswalk.length > 6 && (
                  <div className="relative w-64">
                    <Search className="w-3.5 h-3.5 absolute left-2 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Tìm xã cũ hoặc xã mới..."
                      value={crosswalkSearch}
                      onChange={e => setCrosswalkSearch(e.target.value)}
                      className="w-full text-xs pl-7 pr-2 py-0.5 border border-amber-300 bg-white"
                    />
                  </div>
                )}
              </div>

              <div className="flex flex-wrap gap-1.5 max-h-48 overflow-y-auto p-1 bg-white border border-amber-200">
                {filteredCustomCrosswalk.map((rule, rIdx) => {
                  const actualIdx = customCrosswalk.findIndex(
                    c => c.communeA === rule.communeA && c.communeB === rule.communeB
                  );
                  return (
                    <span
                      key={rIdx}
                      className="bg-amber-50/60 border border-amber-300 text-[11px] px-2 py-0.5 font-medium text-amber-950 flex items-center gap-1.5 hover:bg-amber-100/60 transition-colors"
                    >
                      <span className="font-bold text-sky-800">{rule.communeA}</span>
                      <span className="text-amber-600 font-bold">➔</span>
                      <span className="font-bold text-emerald-800">{rule.communeB}</span>
                      {rule.communeCodeB && (
                        <span className="bg-amber-100 text-amber-900 text-[10px] px-1 font-mono">
                          [{rule.communeCodeB}]
                        </span>
                      )}
                      <button
                        type="button"
                        onClick={() => handleRemoveCustomRule(actualIdx !== -1 ? actualIdx : rIdx)}
                        className="text-red-500 hover:text-red-700 font-bold ml-1 cursor-pointer"
                        title="Xóa quy tắc này"
                      >
                        ✕
                      </button>
                    </span>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        {/* NÚT THỰC HIỆN GHÉP NỐI */}
        <div className="pt-2 flex items-center justify-between border-t border-slate-200">
          <div className="text-[11px] text-slate-600">
            Hệ thống sẽ chạy qua 4 tầng: Khớp CCCD/SĐT ➔ Tự học sơ đồ sáp nhập xã thôn ➔ Khớp Hộ + Thôn mềm ➔ Bóc tách các hộ cần rà soát.
          </div>
          <button
            type="button"
            onClick={handleRunMatching}
            disabled={isProcessing}
            className={`font-bold text-xs px-5 py-2 text-white shadow-sm flex items-center gap-2 cursor-pointer border-0 ${
              isProcessing ? "bg-slate-400 cursor-not-allowed" : "bg-[#286e42] hover:bg-[#205835]"
            }`}
          >
            {isProcessing ? (
              <>
                <RefreshCw className="w-4 h-4 animate-spin" />
                <span>Đang dò tìm &amp; ghép nối sáp nhập...</span>
              </>
            ) : (
              <>
                <Sparkles className="w-4 h-4 text-amber-300" />
                <span>BẮT ĐẦU DÒ TÌM &amp; GHÉP NỐI ĐỊA BÀN MỚI CHO HỘ &gt;</span>
              </>
            )}
          </button>
        </div>
      </div>

      {/* 4. KẾT QUẢ VÀ BÁO CÁO SAU KHI CHẠY XONG */}
      {matchingResult && (
        <div className="mx-4 space-y-3 pb-3">
          {/* THẺ CHỈ SỐ THỐNG KÊ */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2.5">
            <div className="bg-white border border-sky-300 p-2.5">
              <span className="text-[11px] font-bold text-slate-600 block">Tổng số hộ Kỳ 1 (NN cũ):</span>
              <div className="text-xl font-bold font-mono text-sky-900">
                {matchingResult.stats.totalA.toLocaleString("vi-VN")}
              </div>
              <span className="text-[10px] text-slate-500">Đối chiếu với {matchingResult.stats.totalB.toLocaleString("vi-VN")} bản ghi mới</span>
            </div>

            <div className="bg-emerald-50 border-2 border-emerald-400 p-2.5">
              <span className="text-[11px] font-bold text-emerald-900 block flex items-center gap-1">
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                ĐÃ TÌM THẤY ĐỊA BÀN MỚI:
              </span>
              <div className="text-xl font-bold font-mono text-emerald-800">
                {matchingResult.stats.matchedCount.toLocaleString("vi-VN")} hộ
                <span className="text-xs font-normal text-emerald-700 ml-1.5">
                  ({matchingResult.stats.matchRate.toFixed(1)}%)
                </span>
              </div>
              <span className="text-[10px] text-emerald-800">
                {matchingResult.stats.exactIdMatches > 0 && `${matchingResult.stats.exactIdMatches} theo CCCD • `}
                {matchingResult.stats.highMatches} khớp cao • {matchingResult.stats.mediumMatches} khớp TB
              </span>
            </div>

            <div className="bg-amber-50 border border-amber-300 p-2.5">
              <span className="text-[11px] font-bold text-amber-900 block flex items-center gap-1">
                <HelpCircle className="w-3.5 h-3.5 text-amber-600" />
                NGHI NGỜ CẦN XÁC NHẬN:
              </span>
              <div className="text-xl font-bold font-mono text-amber-800">
                {matchingResult.records.filter(r => r.status === "medium_match" || (r.status === "unmatched" && (r.otherCandidates?.length || 0) > 0)).length.toLocaleString("vi-VN")} hộ
              </div>
              <span className="text-[10px] text-amber-800">Trùng tên nhưng thôn sáp nhập tương đối</span>
            </div>

            <div className="bg-rose-50 border border-rose-300 p-2.5">
              <span className="text-[11px] font-bold text-rose-900 block flex items-center gap-1">
                <AlertTriangle className="w-3.5 h-3.5 text-rose-600" />
                CHƯA TÌM THẤY ĐỊA BÀN:
              </span>
              <div className="text-xl font-bold font-mono text-rose-800">
                {matchingResult.stats.unmatchedCount.toLocaleString("vi-VN")} hộ
              </div>
              <span className="text-[10px] text-rose-700">Cần cán bộ kiểm tra thực địa</span>
            </div>
          </div>

          {/* THANH TAB ĐIỀU HƯỚNG */}
          <div className="bg-white border border-slate-300 flex flex-wrap items-center justify-between gap-2 p-1.5">
            <div className="flex flex-wrap items-center gap-1">
              <button
                type="button"
                onClick={() => setActiveTab("matched")}
                className={`text-xs font-bold px-3 py-1.5 flex items-center gap-1.5 cursor-pointer border-0 ${
                  activeTab === "matched"
                    ? "bg-[#286e42] text-white shadow-2xs"
                    : "bg-slate-100 text-slate-700 hover:bg-slate-200"
                }`}
              >
                <CheckCircle2 className="w-3.5 h-3.5" />
                <span>1. ĐÃ GHÉP ĐỊA BÀN MỚI ({matchingResult.stats.matchedCount.toLocaleString("vi-VN")})</span>
              </button>

              <button
                type="button"
                onClick={() => setActiveTab("crosswalk")}
                className={`text-xs font-bold px-3 py-1.5 flex items-center gap-1.5 cursor-pointer border-0 ${
                  activeTab === "crosswalk"
                    ? "bg-indigo-700 text-white shadow-2xs"
                    : "bg-slate-100 text-slate-700 hover:bg-slate-200"
                }`}
              >
                <GitMerge className="w-3.5 h-3.5" />
                <span>2. BẢNG ÁNH XẠ XÃ/THÔN SÁP NHẬP ({matchingResult.communeCrosswalk.length})</span>
              </button>

              <button
                type="button"
                onClick={() => setActiveTab("review")}
                className={`text-xs font-bold px-3 py-1.5 flex items-center gap-1.5 cursor-pointer border-0 ${
                  activeTab === "review"
                    ? "bg-amber-600 text-white shadow-2xs"
                    : "bg-slate-100 text-slate-700 hover:bg-slate-200"
                }`}
              >
                <UserCheck className="w-3.5 h-3.5" />
                <span>3. HỘ NGHI NGỜ (XÁC NHẬN 1 CLICK)</span>
              </button>

              <button
                type="button"
                onClick={() => setActiveTab("unmatched")}
                className={`text-xs font-bold px-3 py-1.5 flex items-center gap-1.5 cursor-pointer border-0 ${
                  activeTab === "unmatched"
                    ? "bg-rose-700 text-white shadow-2xs"
                    : "bg-slate-100 text-slate-700 hover:bg-slate-200"
                }`}
              >
                <AlertTriangle className="w-3.5 h-3.5" />
                <span>4. CHƯA TÌM THẤY ({matchingResult.stats.unmatchedCount.toLocaleString("vi-VN")})</span>
              </button>
            </div>

            {/* BỘ LỌC TÌM KIẾM NHANH */}
            <div className="flex items-center gap-2">
              {uniqueCommunesA.length > 1 && (
                <select
                  value={filterCommune}
                  onChange={e => setFilterCommune(e.target.value)}
                  className="text-xs border border-slate-300 bg-white px-2 py-1 text-slate-800"
                >
                  <option value="ALL">-- Tất cả Xã cũ --</option>
                  {uniqueCommunesA.map(comm => (
                    <option key={comm} value={comm}>
                      Xã: {comm}
                    </option>
                  ))}
                </select>
              )}

              <div className="relative">
                <input
                  type="text"
                  placeholder="Tìm theo tên chủ hộ, thôn, xã..."
                  value={searchQuery}
                  onChange={e => setSearchQuery(e.target.value)}
                  className="text-xs border border-slate-300 bg-white pl-7 pr-2 py-1 w-44 md:w-56"
                />
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2 top-2" />
              </div>
            </div>
          </div>

          {/* TAB 1: DANH SÁCH HỘ ĐÃ GHÉP ĐỊA BÀN MỚI */}
          {activeTab === "matched" && (
            <div className="bg-white border border-slate-300 overflow-x-auto">
              <table className="w-full text-xs text-left border-collapse">
                <thead className="bg-slate-100 text-slate-800 font-bold border-b border-slate-300">
                  <tr>
                    <th className="p-2 border-r border-slate-300 w-12 text-center">STT</th>
                    <th className="p-2 border-r border-slate-300 bg-sky-50 text-sky-950">
                      Chủ Hộ (Kỳ 1 NN cũ)
                    </th>
                    <th className="p-2 border-r border-slate-300 bg-sky-50 text-sky-950">
                      Địa Bàn Cũ (Xã - Thôn)
                    </th>
                    <th className="p-2 border-r border-slate-300 w-8 text-center bg-slate-200">
                      ➔
                    </th>
                    <th className="p-2 border-r border-slate-300 bg-emerald-50 text-emerald-950">
                      Chủ Cơ Sở / Hộ (Kỳ 2 mới)
                    </th>
                    <th className="p-2 border-r border-slate-300 bg-emerald-50 text-emerald-950">
                      Địa Bàn MỚI (Xã - Thôn)
                    </th>
                    <th className="p-2 border-r border-slate-300 bg-emerald-100 text-emerald-950 font-mono">
                      MÃ XÃ MỚI
                    </th>
                    <th className="p-2 border-r border-slate-300 bg-emerald-100 text-emerald-950 font-mono">
                      MÃ THÔN MỚI
                    </th>
                    <th className="p-2 border-r border-slate-300 text-center w-24">Độ Tin Cậy</th>
                    <th className="p-2">Căn Cứ Ghép Nối</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {displayedRecords.length === 0 ? (
                    <tr>
                      <td colSpan={10} className="p-6 text-center text-slate-500 italic">
                        Không có bản ghi nào phù hợp với bộ lọc.
                      </td>
                    </tr>
                  ) : (
                    displayedRecords.slice(0, 100).map((r, idx) => (
                      <tr key={idx} className="hover:bg-amber-50/40">
                        <td className="p-2 border-r border-slate-200 text-center font-mono text-slate-600">
                          {idx + 1}
                        </td>
                        <td className="p-2 border-r border-slate-200 font-bold text-slate-900">
                          {r.householdHeadA}
                          {r.identifierA && (
                            <span className="block text-[10px] font-mono text-slate-500 font-normal">
                              ID: {r.identifierA}
                            </span>
                          )}
                        </td>
                        <td className="p-2 border-r border-slate-200 text-slate-700">
                          <div className="font-medium">{r.communeA || "(Trống xã)"}</div>
                          <div className="text-[11px] text-slate-500">{r.villageA || "(Trống thôn)"}</div>
                        </td>
                        <td className="p-2 border-r border-slate-200 text-center text-slate-400 font-bold bg-slate-50">
                          ➔
                        </td>
                        <td className="p-2 border-r border-slate-200 font-bold text-emerald-950">
                          {r.householdHeadB}
                          {r.identifierB && (
                            <span className="block text-[10px] font-mono text-slate-500 font-normal">
                              ID: {r.identifierB}
                            </span>
                          )}
                        </td>
                        <td className="p-2 border-r border-slate-200 text-emerald-900 font-medium">
                          <div>{r.communeB || "(Chưa có tên xã mới)"}</div>
                          <div className="text-[11px] text-emerald-700">{r.villageB || "(Chưa có tên thôn mới)"}</div>
                        </td>
                        <td className="p-2 border-r border-slate-200 font-mono font-bold text-emerald-900 bg-emerald-50/60">
                          {r.communeCodeB || "--"}
                        </td>
                        <td className="p-2 border-r border-slate-200 font-mono font-bold text-emerald-900 bg-emerald-50/60">
                          {r.villageCodeB || "--"}
                        </td>
                        <td className="p-2 border-r border-slate-200 text-center">
                          <span
                            className={`inline-block px-1.5 py-0.5 text-[10px] font-bold font-mono ${
                              r.confidence >= 95
                                ? "bg-emerald-100 text-emerald-900 border border-emerald-300"
                                : r.confidence >= 80
                                ? "bg-sky-100 text-sky-900 border border-sky-300"
                                : "bg-amber-100 text-amber-900 border border-amber-300"
                            }`}
                          >
                            {r.confidence}%
                          </span>
                        </td>
                        <td className="p-2 text-slate-600 text-[11px] leading-tight">
                          {r.matchReason}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>

              {displayedRecords.length > 100 && (
                <div className="p-2 bg-slate-50 text-center text-xs text-slate-600 border-t border-slate-200">
                  Đang hiển thị 100 dòng đầu tiên trên tổng số {displayedRecords.length.toLocaleString("vi-VN")} dòng. Xuất Excel để xem toàn bộ.
                </div>
              )}
            </div>
          )}

          {/* TAB 2: BẢNG ÁNH XẠ XÃ / THÔN SÁP NHẬP */}
          {activeTab === "crosswalk" && (
            <div className="space-y-4">
              {/* Bảng Xã sáp nhập */}
              <div className="bg-white border border-slate-300 p-3 space-y-2">
                <div className="flex items-center justify-between">
                  <h4 className="text-xs font-bold text-indigo-950 uppercase flex items-center gap-1.5">
                    <Building2 className="w-4 h-4 text-indigo-700" />
                    BẢNG ÁNH XẠ SÁP NHẬP CẤP XÃ (QUY ĐỔI MÃ XÃ CŨ ➔ MÃ XÃ MỚI)
                  </h4>
                  <span className="text-[11px] text-slate-600">
                    Phát hiện {matchingResult.communeCrosswalk.length} quan hệ sáp nhập
                  </span>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left border-collapse">
                    <thead className="bg-indigo-50 text-indigo-950 font-bold border-b border-indigo-200">
                      <tr>
                        <th className="p-2 border-r border-indigo-200 w-12 text-center">STT</th>
                        <th className="p-2 border-r border-indigo-200">Tên Xã Cũ (Kỳ NN cũ)</th>
                        <th className="p-2 border-r border-indigo-200 font-mono">Mã Xã Cũ</th>
                        <th className="p-2 border-r border-indigo-200 text-center w-12 font-bold text-indigo-600">
                          ➔
                        </th>
                        <th className="p-2 border-r border-indigo-200 font-bold text-emerald-950">
                          SÁP NHẬP VÀO XÃ MỚI (Kỳ Cá thể)
                        </th>
                        <th className="p-2 border-r border-indigo-200 font-mono font-bold text-emerald-950">
                          MÃ XÃ MỚI
                        </th>
                        <th className="p-2 border-r border-indigo-200 text-right">Số Hộ Dân Khớp</th>
                        <th className="p-2 text-center">Độ Tin Cậy</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {matchingResult.communeCrosswalk.map((item, idx) => (
                        <tr key={idx} className="hover:bg-indigo-50/40">
                          <td className="p-2 border-r border-slate-200 text-center font-mono">{idx + 1}</td>
                          <td className="p-2 border-r border-slate-200 font-bold text-sky-950">
                            {item.communeNameA}
                          </td>
                          <td className="p-2 border-r border-slate-200 font-mono text-slate-600">
                            {item.communeCodeA || "--"}
                          </td>
                          <td className="p-2 border-r border-slate-200 text-center text-indigo-500 font-bold bg-slate-50">
                            ➔
                          </td>
                          <td className="p-2 border-r border-slate-200 font-bold text-emerald-900">
                            {item.communeNameB}
                          </td>
                          <td className="p-2 border-r border-slate-200 font-mono font-bold text-emerald-900 bg-emerald-50/50">
                            {item.communeCodeB || "--"}
                          </td>
                          <td className="p-2 border-r border-slate-200 text-right font-mono font-bold text-slate-800">
                            {item.matchedCount.toLocaleString("vi-VN")} hộ
                          </td>
                          <td className="p-2 text-center">
                            <span className="inline-block px-2 py-0.5 bg-emerald-100 text-emerald-900 text-[10px] font-bold">
                              {item.confidence}%
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Bảng Thôn sáp nhập */}
              <div className="bg-white border border-slate-300 p-3 space-y-2">
                <div className="flex items-center justify-between">
                  <h4 className="text-xs font-bold text-sky-950 uppercase flex items-center gap-1.5">
                    <MapPin className="w-4 h-4 text-sky-700" />
                    BẢNG ÁNH XẠ SÁP NHẬP CẤP THÔN / ĐỊA BÀN
                  </h4>
                  <span className="text-[11px] text-slate-600">
                    Phát hiện {matchingResult.villageCrosswalk.length} quan hệ thôn/địa bàn
                  </span>
                </div>

                <div className="overflow-x-auto max-h-80 overflow-y-auto">
                  <table className="w-full text-xs text-left border-collapse">
                    <thead className="bg-sky-50 text-sky-950 font-bold border-b border-sky-200 sticky top-0">
                      <tr>
                        <th className="p-2 border-r border-sky-200 w-12 text-center">STT</th>
                        <th className="p-2 border-r border-sky-200">Xã Cũ</th>
                        <th className="p-2 border-r border-sky-200">Thôn Cũ (Kỳ 1)</th>
                        <th className="p-2 border-r border-sky-200 font-mono">Mã Thôn Cũ</th>
                        <th className="p-2 border-r border-sky-200 text-center w-12 font-bold text-sky-600">➔</th>
                        <th className="p-2 border-r border-sky-200">Xã Mới</th>
                        <th className="p-2 border-r border-sky-200 font-bold text-emerald-950">
                          Thôn / Địa Bàn Mới (Kỳ 2)
                        </th>
                        <th className="p-2 border-r border-sky-200 font-mono font-bold text-emerald-950">
                          MÃ THÔN MỚI
                        </th>
                        <th className="p-2 text-right">Số Hộ Dân Khớp</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {matchingResult.villageCrosswalk.map((item, idx) => (
                        <tr key={idx} className="hover:bg-sky-50/40">
                          <td className="p-2 border-r border-slate-200 text-center font-mono">{idx + 1}</td>
                          <td className="p-2 border-r border-slate-200 text-slate-700">{item.communeNameA}</td>
                          <td className="p-2 border-r border-slate-200 font-bold text-sky-950">{item.villageNameA}</td>
                          <td className="p-2 border-r border-slate-200 font-mono text-slate-600">{item.villageCodeA || "--"}</td>
                          <td className="p-2 border-r border-slate-200 text-center text-sky-500 font-bold bg-slate-50">➔</td>
                          <td className="p-2 border-r border-slate-200 text-slate-700">{item.communeNameB}</td>
                          <td className="p-2 border-r border-slate-200 font-bold text-emerald-900">{item.villageNameB}</td>
                          <td className="p-2 border-r border-slate-200 font-mono font-bold text-emerald-900 bg-emerald-50/50">{item.villageCodeB || "--"}</td>
                          <td className="p-2 text-right font-mono font-bold text-slate-800">{item.matchedCount.toLocaleString("vi-VN")} hộ</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* TAB 3: HỘ NGHI NGỜ & XÁC NHẬN THỦ CÔNG 1 CLICK */}
          {activeTab === "review" && (
            <div className="bg-white border border-slate-300 p-3 space-y-3">
              <div className="flex items-center justify-between pb-2 border-b border-slate-200">
                <div>
                  <h4 className="text-xs font-bold text-amber-950 uppercase flex items-center gap-1.5">
                    <UserCheck className="w-4 h-4 text-amber-600" />
                    DANH SÁCH HỘ NGHI NGỜ (HỖ TRỢ CÁN BỘ CHỐNG TRÙNG TÊN &amp; CHỌN NHANH)
                  </h4>
                  <p className="text-[11px] text-slate-600">
                    Hệ thống liệt kê các hộ có độ tương đồng từ 70% đến 84% hoặc có nhiều người cùng tên trên địa bàn sáp nhập. Bạn chỉ cần bấm nút <b>"Chọn hộ này"</b> để chốt.
                  </p>
                </div>
              </div>

              <div className="space-y-2.5">
                {displayedRecords.length === 0 ? (
                  <div className="p-8 text-center text-slate-500 italic bg-slate-50">
                    Không có hộ nào trong nhóm nghi ngờ cần rà soát thủ công.
                  </div>
                ) : (
                  displayedRecords.map((r, idx) => (
                    <div key={idx} className="border border-amber-300 bg-amber-50/30 p-3 space-y-2">
                      <div className="flex flex-col md:flex-row md:items-center justify-between gap-2 pb-1.5 border-b border-amber-200">
                        <div className="flex items-center gap-2">
                          <span className="font-mono text-xs font-bold bg-amber-200 text-amber-950 px-1.5 py-0.5">
                            #{r.indexA + 1}
                          </span>
                          <span className="font-bold text-slate-900 text-xs">
                            {r.householdHeadA}
                          </span>
                          <span className="text-xs text-slate-600">
                            • Xã cũ: <b className="text-sky-900">{r.communeA || "Trống"}</b> • Thôn cũ: <b className="text-sky-900">{r.villageA || "Trống"}</b>
                          </span>
                          {r.identifierA && (
                            <span className="text-[11px] font-mono bg-white px-1 border border-slate-300">
                              CCCD: {r.identifierA}
                            </span>
                          )}
                        </div>

                        <div className="text-[11px] text-amber-900 font-medium">
                          {r.matchReason}
                        </div>
                      </div>

                      {/* Danh sách các ứng viên có thể chọn */}
                      <div className="space-y-1.5">
                        <span className="text-[10px] font-bold text-slate-700 uppercase block">
                          Gợi ý các cơ sở/hộ tương đồng trên địa bàn mới:
                        </span>

                        {/* Ứng viên hiện tại nếu có */}
                        {r.rowB && (
                          <div className="bg-emerald-50 border border-emerald-300 p-2 flex items-center justify-between gap-2">
                            <div className="text-xs">
                              <span className="font-bold text-emerald-950">{r.householdHeadB}</span>
                              <span className="text-slate-600 ml-2">
                                Xã mới: <b>{r.communeB}</b> (Mã: <code className="text-emerald-800">{r.communeCodeB || "--"}</code>) • Thôn mới: <b>{r.villageB}</b> (Mã: <code className="text-emerald-800">{r.villageCodeB || "--"}</code>)
                              </span>
                            </div>
                            <span className="text-xs font-bold text-emerald-700 bg-emerald-100 px-2 py-0.5 border border-emerald-300">
                              ✓ Đang chọn ({r.confidence}%)
                            </span>
                          </div>
                        )}

                        {/* Các ứng viên khác */}
                        {r.otherCandidates && r.otherCandidates.map((cand, cIdx) => (
                          <div key={cIdx} className="bg-white border border-slate-300 p-2 flex items-center justify-between gap-2 hover:bg-slate-50">
                            <div className="text-xs">
                              <span className="font-bold text-slate-900">{cand.householdHeadB}</span>
                              <span className="text-slate-600 ml-2">
                                Xã mới: <b>{cand.communeB}</b> • Thôn mới: <b>{cand.villageB}</b>
                              </span>
                            </div>
                            <button
                              type="button"
                              onClick={() => handleSelectCandidateManually(r.indexA, cand.rowB)}
                              className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-[11px] px-2.5 py-1 cursor-pointer border-0 shadow-2xs"
                            >
                              Chọn Hộ Này ({cand.confidence}%)
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          )}

          {/* TAB 4: CHƯA TÌM THẤY TRÊN ĐỊA BÀN MỚI */}
          {activeTab === "unmatched" && (
            <div className="bg-white border border-rose-300 p-3 space-y-3">
              <div className="flex items-center justify-between pb-2 border-b border-rose-200">
                <div>
                  <h4 className="text-xs font-bold text-rose-950 uppercase flex items-center gap-1.5">
                    <AlertTriangle className="w-4 h-4 text-rose-600" />
                    DANH SÁCH CÁC HỘ CHƯA TÌM THẤY TRÊN ĐỊA BÀN MỚI ({displayedRecords.length.toLocaleString("vi-VN")} HỘ)
                  </h4>
                  <p className="text-[11px] text-slate-600">
                    Các hộ này không tìm thấy tên tương đồng trên địa bàn mới sau sáp nhập. Có thể hộ đã chuyển đi, nghỉ kinh doanh hoặc sáp nhập sang đơn vị hành chính ngoài phạm vi tệp.
                  </p>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left border-collapse">
                  <thead className="bg-rose-50 text-rose-950 font-bold border-b border-rose-200">
                    <tr>
                      <th className="p-2 border-r border-rose-200 w-12 text-center">STT</th>
                      <th className="p-2 border-r border-rose-200">Tên Chủ Hộ (Kỳ 1)</th>
                      <th className="p-2 border-r border-rose-200">Xã Cũ</th>
                      <th className="p-2 border-r border-rose-200 font-mono">Mã Xã Cũ</th>
                      <th className="p-2 border-r border-rose-200">Thôn Cũ</th>
                      <th className="p-2 border-r border-rose-200 font-mono">Mã Thôn Cũ</th>
                      <th className="p-2 border-r border-rose-200">CCCD/SĐT</th>
                      <th className="p-2">Ghi Chú Rà Soát</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {displayedRecords.slice(0, 100).map((r, idx) => (
                      <tr key={idx} className="hover:bg-rose-50/40">
                        <td className="p-2 border-r border-slate-200 text-center font-mono">{idx + 1}</td>
                        <td className="p-2 border-r border-slate-200 font-bold text-slate-900">{r.householdHeadA}</td>
                        <td className="p-2 border-r border-slate-200 text-slate-700">{r.communeA}</td>
                        <td className="p-2 border-r border-slate-200 font-mono text-slate-600">{r.communeCodeA || "--"}</td>
                        <td className="p-2 border-r border-slate-200 text-slate-700">{r.villageA}</td>
                        <td className="p-2 border-r border-slate-200 font-mono text-slate-600">{r.villageCodeA || "--"}</td>
                        <td className="p-2 border-r border-slate-200 font-mono text-slate-700">{r.identifierA || "--"}</td>
                        <td className="p-2 text-rose-800 text-[11px]">
                          <div>{r.matchReason}</div>
                          {r.communeB && (
                            <div className="text-[10px] text-sky-700 font-semibold mt-0.5">
                              ➔ Đã định vị sáp nhập vào Xã mới: <span className="font-bold">{r.communeB}</span>
                              {r.communeCodeB && ` [Mã: ${r.communeCodeB}]`}
                            </div>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}

      {/* MODAL DÁN NHANH TỪ EXCEL */}
      {showPasteModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-300 shadow-2xl max-w-2xl w-full flex flex-col max-h-[90vh]">
            <div className="bg-[#286e42] text-white px-4 py-2.5 flex items-center justify-between">
              <div className="flex items-center gap-2 font-bold text-sm">
                <ClipboardPaste className="w-4 h-4" />
                Dán nhanh bảng quy đổi Xã cũ ➔ Xã mới từ Excel
              </div>
              <button
                type="button"
                onClick={() => setShowPasteModal(false)}
                className="text-white/80 hover:text-white cursor-pointer border-0 bg-transparent"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-4 space-y-3 overflow-y-auto flex-1">
              <div className="bg-sky-50 border border-sky-200 p-2.5 text-xs text-sky-900 space-y-1">
                <div className="font-bold flex items-center gap-1 text-sky-950">
                  <CheckCircle2 className="w-4 h-4 text-sky-600" /> Hướng dẫn dán từ Excel:
                </div>
                <div>
                  1. Trong file Excel bảng sáp nhập của bạn, bôi đen 2 cột: <strong>Cột 1: Tên Xã cũ</strong> và <strong>Cột 2: Tên Xã mới</strong> (có thể thêm Cột 3: Mã xã mới nếu có).
                </div>
                <div>
                  2. Nhấn <strong>Ctrl + C</strong> trên Excel để sao chép.
                </div>
                <div>
                  3. Bấm vào khung bên dưới và nhấn <strong>Ctrl + V</strong> (Dán), sau đó bấm nút &ldquo;Nạp quy tắc vào hệ thống&rdquo;.
                </div>
              </div>

              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="text-xs font-bold text-slate-700">
                    Nội dung sao chép từ Excel:
                  </label>
                  <button
                    type="button"
                    onClick={() => {
                      const sampleText = SAMPLE_24_COMMUNES.map(c => `${c.communeA}\t${c.communeB}`).join("\n");
                      setPasteRawText(sampleText);
                    }}
                    className="text-[11px] text-sky-700 hover:text-sky-900 font-semibold cursor-pointer underline border-0 bg-transparent"
                  >
                    Dán thử 24 xã mẫu trong ảnh của bạn
                  </button>
                </div>
                <textarea
                  rows={10}
                  value={pasteRawText}
                  onChange={e => setPasteRawText(e.target.value)}
                  placeholder={`Ví dụ định dạng copy từ Excel:\nXã An Đồng\tA Sào\nXã An Hiệp\tA Sào\nXã An Thái\tA Sào\nXã An Khê\tA Sào\nXã Tây Giang\tÁi Quốc\n...`}
                  className="w-full font-mono text-xs border border-slate-300 p-2.5 bg-slate-50 focus:bg-white focus:outline-none focus:border-sky-600"
                />
              </div>

              {pasteRawText.trim() && (
                <div className="text-[11px] text-slate-600">
                  Ước tính nhận diện: <strong>{pasteRawText.split(/\r?\n/).filter(l => l.trim()).length} dòng</strong>.
                </div>
              )}
            </div>

            <div className="bg-slate-100 px-4 py-2.5 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={() => {
                  setPasteRawText("");
                  setShowPasteModal(false);
                }}
                className="text-xs px-3 py-1.5 text-slate-700 hover:bg-slate-200 border border-slate-300 font-medium cursor-pointer"
              >
                Hủy bỏ
              </button>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setPasteRawText("")}
                  disabled={!pasteRawText}
                  className="text-xs px-2.5 py-1.5 text-slate-600 hover:text-slate-900 font-medium cursor-pointer border-0 bg-transparent disabled:opacity-40"
                >
                  Xóa trắng khung
                </button>
                <button
                  type="button"
                  onClick={handleApplyPastedExcel}
                  disabled={!pasteRawText.trim()}
                  className="text-xs px-4 py-1.5 bg-[#286e42] hover:bg-[#205835] disabled:bg-slate-300 text-white font-bold flex items-center gap-1.5 cursor-pointer shadow-xs border-0"
                >
                  <Check className="w-3.5 h-3.5" />
                  Nạp quy tắc vào hệ thống
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
