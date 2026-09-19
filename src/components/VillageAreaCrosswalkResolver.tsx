import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  MapPin,
  GitMerge,
  Search,
  CheckCircle2,
  AlertTriangle,
  HelpCircle,
  Download,
  ArrowRight,
  Sparkles,
  RefreshCw,
  FileSpreadsheet,
  Upload,
  Check,
  X,
  Filter,
  Eye,
  SlidersHorizontal,
  ChevronDown,
  BookOpen,
  Database,
  ExternalLink,
  Users,
  Archive,
  FolderCheck,
  RotateCcw,
  Building2,
  Layers,
  CheckSquare
} from "lucide-react";
import * as XLSX from "xlsx";
import { HouseholdClusterResolver } from "./HouseholdClusterResolver";
import { CommuneClusterMatrixModal } from "./CommuneClusterMatrixModal";
import {
  executeVillageAreaCrosswalk,
  VillageAreaCrosswalkConfig,
  VillageMatchResult,
  cleanVillageName,
  parseHouseholdCount
} from "../utils/villageAreaCrosswalk";
import {
  OFFICIAL_COMMUNE_MERGE_DATA,
  findOfficialNewCommune,
  findOfficialByCommuneCode,
  findOfficialNewCommuneCode
} from "../data/officialCommuneCrosswalk";
import { parse2DArrayWithSmartHeader } from "../utils/sharedHelpers";

// Dữ liệu mẫu thực tế trích xuất trực tiếp từ 2 ảnh bảng kê của người dùng
export const SAMPLE_OLD_VILLAGES_FROM_IMAGE = [
  { STT: 5, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11953, "Tên xã": "Phường Hiến Nam", "Mã Thôn, Tổ dân phố": "001", "Tên thôn, Tổ dân phố": "Tổ dân phố Chùa Chuông", "Tổng số hộ của thôn, Tổ dân phố": 623 },
  { STT: 6, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11953, "Tên xã": "Phường Hiến Nam", "Mã Thôn, Tổ dân phố": "002", "Tên thôn, Tổ dân phố": "Tổ dân phố Nhân Dục", "Tổng số hộ của thôn, Tổ dân phố": 836 },
  { STT: 7, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11953, "Tên xã": "Phường Hiến Nam", "Mã Thôn, Tổ dân phố": "003", "Tên thôn, Tổ dân phố": "TDP Ấp Dâu", "Tổng số hộ của thôn, Tổ dân phố": 545 },
  { STT: 8, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11953, "Tên xã": "Phường Hiến Nam", "Mã Thôn, Tổ dân phố": "004", "Tên thôn, Tổ dân phố": "Tổ dân phố An Vũ", "Tổng số hộ của thôn, Tổ dân phố": 429 },
  { STT: 9, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11953, "Tên xã": "Phường Hiến Nam", "Mã Thôn, Tổ dân phố": "005", "Tên thôn, Tổ dân phố": "TDP An Thịnh", "Tổng số hộ của thôn, Tổ dân phố": 352 },
  { STT: 10, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11953, "Tên xã": "Phường Hiến Nam", "Mã Thôn, Tổ dân phố": "006", "Tên thôn, Tổ dân phố": "Tổ dân phố An Đông", "Tổng số hộ của thôn, Tổ dân phố": 427 },
  { STT: 12, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11956, "Tên xã": "Phường An Tảo", "Mã Thôn, Tổ dân phố": "001", "Tên thôn, Tổ dân phố": "Tổ dân phố An Thượng", "Tổng số hộ của thôn, Tổ dân phố": 875 },
  { STT: 13, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11956, "Tên xã": "Phường An Tảo", "Mã Thôn, Tổ dân phố": "002", "Tên thôn, Tổ dân phố": "Tổ dân phố An Bình", "Tổng số hộ của thôn, Tổ dân phố": 960 },
  { STT: 14, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11956, "Tên xã": "Phường An Tảo", "Mã Thôn, Tổ dân phố": "003", "Tên thôn, Tổ dân phố": "Tổ dân phố An Dương", "Tổng số hộ của thôn, Tổ dân phố": 862 },
  { STT: 15, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11956, "Tên xã": "Phường An Tảo", "Mã Thôn, Tổ dân phố": "004", "Tên thôn, Tổ dân phố": "Tổ dân phố Chợ Gạo", "Tổng số hộ của thôn, Tổ dân phố": 556 },
  { STT: 16, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11959, "Tên xã": "Phường Lê Lợi", "Mã Thôn, Tổ dân phố": "001", "Tên thôn, Tổ dân phố": "Điện Biên", "Tổng số hộ của thôn, Tổ dân phố": 468 },
  { STT: 18, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11959, "Tên xã": "Phường Lê Lợi", "Mã Thôn, Tổ dân phố": "003", "Tên thôn, Tổ dân phố": "Hai Bà Trưng", "Tổng số hộ của thôn, Tổ dân phố": 300 },
  { STT: 23, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11959, "Tên xã": "Phường Lê Lợi", "Mã Thôn, Tổ dân phố": "008", "Tên thôn, Tổ dân phố": "Điện Biên 1", "Tổng số hộ của thôn, Tổ dân phố": 174 },
  { STT: 24, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11959, "Tên xã": "Phường Lê Lợi", "Mã Thôn, Tổ dân phố": "009", "Tên thôn, Tổ dân phố": "Điện Biên 2", "Tổng số hộ của thôn, Tổ dân phố": 295 },
  { STT: 25, "Mã tỉnh": 33, "Mã huyện": 323, "Mã xã": 11959, "Tên xã": "Phường Lê Lợi", "Mã Thôn, Tổ dân phố": "010", "Tên thôn, Tổ dân phố": "Bãi Sậy", "Tổng số hộ của thôn, Tổ dân phố": 160 },
];

export const SAMPLE_NEW_VILLAGES_FROM_IMAGE = [
  { STT: 1, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Chợ Gạo", "Mã thôn/TDP": 1, "Tổng số hộ của thôn/TDP": 556, "Tên ĐBĐT": "Chợ Gạo_001", "Mã ĐBĐT": 1, "Tổng số hộ của ĐBĐT": 278, "Địa bàn mẫu": "x" },
  { STT: 3, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "An Bình", "Mã thôn/TDP": 2, "Tổng số hộ của thôn/TDP": 992, "Tên ĐBĐT": "An Bình_001", "Mã ĐBĐT": 3, "Tổng số hộ của ĐBĐT": 496, "Địa bàn mẫu": "x" },
  { STT: 5, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "An Dương", "Mã thôn/TDP": 3, "Tổng số hộ của thôn/TDP": 862, "Tên ĐBĐT": "An Dương_001", "Mã ĐBĐT": 5, "Tổng số hộ của ĐBĐT": 431, "Địa bàn mẫu": "x" },
  { STT: 7, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "An Thượng", "Mã thôn/TDP": 4, "Tổng số hộ của thôn/TDP": 795, "Tên ĐBĐT": "An Thượng_001", "Mã ĐBĐT": 7, "Tổng số hộ của ĐBĐT": 398, "Địa bàn mẫu": "x" },
  { STT: 9, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Chùa Chuông", "Mã thôn/TDP": 5, "Tổng số hộ của thôn/TDP": 623, "Tên ĐBĐT": "Chùa Chuông_001", "Mã ĐBĐT": 9, "Tổng số hộ của ĐBĐT": 312, "Địa bàn mẫu": "x" },
  { STT: 11, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Nhân Dục", "Mã thôn/TDP": 6, "Tổng số hộ của thôn/TDP": 879, "Tên ĐBĐT": "Nhân Dục_001", "Mã ĐBĐT": 11, "Tổng số hộ của ĐBĐT": 440, "Địa bàn mẫu": "x" },
  { STT: 13, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Ấp Dâu", "Mã thôn/TDP": 7, "Tổng số hộ của thôn/TDP": 520, "Tên ĐBĐT": "Ấp Dâu_001", "Mã ĐBĐT": 13, "Tổng số hộ của ĐBĐT": 260, "Địa bàn mẫu": "x" },
  { STT: 15, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "An Thịnh", "Mã thôn/TDP": 8, "Tổng số hộ của thôn/TDP": 351, "Tên ĐBĐT": "An Thịnh", "Mã ĐBĐT": 15, "Tổng số hộ của ĐBĐT": 351, "Địa bàn mẫu": "x" },
  { STT: 17, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "An Đông", "Mã thôn/TDP": 10, "Tổng số hộ của thôn/TDP": 426, "Tên ĐBĐT": "An Đông", "Mã ĐBĐT": 17, "Tổng số hộ của ĐBĐT": 426, "Địa bàn mẫu": "x" },
  { STT: 21, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Bắc Thành", "Mã thôn/TDP": 14, "Tổng số hộ của thôn/TDP": 390, "Tên ĐBĐT": "Bắc Thành", "Mã ĐBĐT": 21, "Tổng số hộ của ĐBĐT": 390, "Địa bàn mẫu": "x" },
  { STT: 23, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Bãi Sậy", "Mã thôn/TDP": 16, "Tổng số hộ của thôn/TDP": 170, "Tên ĐBĐT": "Bãi Sậy", "Mã ĐBĐT": 23, "Tổng số hộ của ĐBĐT": 170, "Địa bàn mẫu": "x" },
  { STT: 25, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Điện Biên II", "Mã thôn/TDP": 18, "Tổng số hộ của thôn/TDP": 410, "Tên ĐBĐT": "Điện Biên II", "Mã ĐBĐT": 25, "Tổng số hộ của ĐBĐT": 410, "Địa bàn mẫu": "x" },
  { STT: 27, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Điện Biên", "Mã thôn/TDP": 20, "Tổng số hộ của thôn/TDP": 350, "Tên ĐBĐT": "Điện Biên", "Mã ĐBĐT": 27, "Tổng số hộ của ĐBĐT": 350, "Địa bàn mẫu": "x" },
  { STT: 29, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Hai Bà Trưng", "Mã thôn/TDP": 22, "Tổng số hộ của thôn/TDP": 255, "Tên ĐBĐT": "Hai Bà Trưng", "Mã ĐBĐT": 29, "Tổng số hộ của ĐBĐT": 255, "Địa bàn mẫu": "x" },
  { STT: 31, "Mã tỉnh": 33, "Tên Thống kê cơ sở": "Thống kê cơ sở Phố Hiến", "Mã Thống kê cơ sở": "K133T01", "Tên xã": "Phường Phố Hiến", "Mã xã": 11953, TTNT: 1, "Tên thôn/TDP": "Lê Lai", "Mã thôn/TDP": 24, "Tổng số hộ của thôn/TDP": 202, "Tên ĐBĐT": "Lê Lai", "Mã ĐBĐT": 31, "Tổng số hộ của ĐBĐT": 202, "Địa bàn mẫu": "x" },
];

interface VillageAreaCrosswalkResolverProps {
  initialDataA?: any[];
  initialColumnsA?: string[];
  initialDataB?: any[];
  initialColumnsB?: string[];
  onApplyMergedToMain?: (mergedData: any[], mergedCols: string[]) => void;
  onClose?: () => void;
}

export const VillageAreaCrosswalkResolver: React.FC<VillageAreaCrosswalkResolverProps> = ({
  initialDataA = [],
  initialColumnsA = [],
  initialDataB = [],
  initialColumnsB = [],
  onApplyMergedToMain,
  onClose,
}) => {
  // Dữ liệu Tệp A (Địa bàn cũ)
  const [dataA, setDataA] = useState<any[]>(initialDataA);
  const [columnsA, setColumnsA] = useState<string[]>(initialColumnsA);
  const [fileNameA, setFileNameA] = useState<string>("Tệp Địa Bàn Cũ (Chưa sáp nhập)");

  // Dữ liệu Tệp B (Địa bàn mới năm nay)
  const [dataB, setDataB] = useState<any[]>(initialDataB);
  const [columnsB, setColumnsB] = useState<string[]>(initialColumnsB);
  const [fileNameB, setFileNameB] = useState<string>("Tệp Địa Bàn Mới (Đã sáp nhập)");

  // Cột Tệp A
  const [colCommuneNameA, setColCommuneNameA] = useState<string>("");
  const [colCommuneCodeA, setColCommuneCodeA] = useState<string>("");
  const [colVillageNameA, setColVillageNameA] = useState<string>("");
  const [colVillageCodeA, setColVillageCodeA] = useState<string>("");
  const [colHouseholdsA, setColHouseholdsA] = useState<string>("");

  // Cột Tệp B
  const [colCommuneNameB, setColCommuneNameB] = useState<string>("");
  const [colCommuneCodeB, setColCommuneCodeB] = useState<string>("");
  const [colVillageNameB, setColVillageNameB] = useState<string>("");
  const [colVillageCodeB, setColVillageCodeB] = useState<string>("");
  const [colHouseholdsB, setColHouseholdsB] = useState<string>("");
  const [colDbdtNameB, setColDbdtNameB] = useState<string>("");
  const [colDbdtCodeB, setColDbdtCodeB] = useState<string>("");
  const [colDbdtHouseholdsB, setColDbdtHouseholdsB] = useState<string>("");

  // Bảng quy đổi xã tùy biến bổ sung
  const [customCommuneCrosswalk, setCustomCommuneCrosswalk] = useState<{ communeA: string; communeB: string }[]>([]);

  // Kết quả đối soát
  const [matchingResults, setMatchingResults] = useState<VillageMatchResult[]>([]);
  const [matchingSummary, setMatchingSummary] = useState<any>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [appliedMsg, setAppliedMsg] = useState<string | null>(null);

  // Phạm vi đối soát theo xã (Chạy theo từng xã & nội bộ xã)
  const [selectedCommuneA, setSelectedCommuneA] = useState<string>("ALL");
  const [strictCommuneScoped, setStrictCommuneScoped] = useState<boolean>(true);

  // Kho Lưu Trữ Bóc Tách: Cất các địa bàn đã khớp đúng đi, chỉ giữ lại địa bàn chưa khớp để làm dần
  const [stashedMatches, setStashedMatches] = useState<VillageMatchResult[]>([]);
  const [activeViewTab, setActiveViewTab] = useState<"ACTIVE" | "STASHED">("ACTIVE");

  // Modal tra cứu 310 xã sáp nhập
  const [showOfficialModal, setShowOfficialModal] = useState<boolean>(false);
  const [officialSearchQuery, setOfficialSearchQuery] = useState<string>("");

  // Chế độ mở Đối Soát Cụm Hộ Dân (≥ 30 Người Trùng Tên)
  const [showClusterResolver, setShowClusterResolver] = useState<boolean>(false);

  // Chế độ mở Ma Trận Gom 65 Xã Mới ➔ Đủ Toàn Bộ Xã Cũ & Địa Bàn
  const [showCommuneClusterMatrix, setShowCommuneClusterMatrix] = useState<boolean>(false);

  // Bộ lọc kết quả hiển thị
  const [filterTab, setFilterTab] = useState<"ALL" | "PERFECT" | "VERY_HIGH" | "HIGH" | "MEDIUM" | "UNMATCHED">("ALL");
  const [searchTerm, setSearchTerm] = useState<string>("");
  const [selectedCandidateRow, setSelectedCandidateRow] = useState<{ index: number; item: VillageMatchResult } | null>(null);

  const fileInputRefA = useRef<HTMLInputElement>(null);
  const fileInputRefB = useRef<HTMLInputElement>(null);

  // Tự động nhận diện cột thông minh & chuẩn xác (phân biệt rạch ròi giữa Tên và Mã)
  const autoDetectColsA = (cols: string[]) => {
    let detectedVillageName = "";
    let detectedVillageCode = "";
    let detectedHouseholds = "";
    let detectedCommuneName = "";
    let detectedCommuneCode = "";

    // Ưu tiên 1: Cột có chữ 'tên' rõ ràng
    for (const c of cols) {
      const lower = c.toLowerCase().trim();
      const hasMa = lower.includes("mã") || lower.includes("code") || lower.startsWith("ma_") || lower.startsWith("ma ");
      const hasTen = lower.includes("tên") || lower.includes("ten");

      // Cột Tên thôn / TDP (Phải có 'tên' và 'thôn'/'tổ dân phố', hoặc không chứa từ 'mã')
      if (!detectedVillageName) {
        if (hasTen && (lower.includes("thôn") || lower.includes("tổ dân phố") || lower.includes("tdp") || lower.includes("địa bàn") || lower.includes("đb"))) {
          detectedVillageName = c;
        }
      }

      // Cột Mã thôn / TDP (Phải có chữ 'mã' và 'thôn'/'tổ dân phố'/'tdp')
      if (!detectedVillageCode) {
        if (hasMa && (lower.includes("thôn") || lower.includes("tổ dân phố") || lower.includes("tdp") || lower.includes("địa bàn") || lower.includes("đb"))) {
          detectedVillageCode = c;
        }
      }

      // Cột Số hộ
      if (!detectedHouseholds) {
        if (lower.includes("tổng số hộ") || lower.includes("số hộ của thôn") || lower.includes("số hộ")) {
          detectedHouseholds = c;
        }
      }

      // Cột Tên xã (Phải có chữ 'tên' và 'xã'/'phường', hoặc là 'xã' mà không có chữ 'mã')
      if (!detectedCommuneName) {
        if ((hasTen && (lower.includes("xã") || lower.includes("phường"))) || (!hasMa && (lower === "xã" || lower === "phường" || lower === "tên xã" || lower === "tên phường" || lower === "xã/phường"))) {
          detectedCommuneName = c;
        }
      }

      // Cột Mã xã
      if (!detectedCommuneCode) {
        if (hasMa && (lower.includes("xã") || lower.includes("phường") || lower.includes("maxa"))) {
          detectedCommuneCode = c;
        }
      }
    }

    // Ưu tiên 2: Tìm bù nếu chưa thấy cột tên thôn (lấy cột nào có chữ 'thôn'/'tổ dân phố' nhưng không có chữ 'mã')
    if (!detectedVillageName) {
      for (const c of cols) {
        const lower = c.toLowerCase().trim();
        if (!lower.includes("mã") && (lower.includes("thôn") || lower.includes("tổ dân phố") || lower.includes("tdp"))) {
          detectedVillageName = c;
          break;
        }
      }
    }

    // Cập nhật state nếu tìm thấy
    if (detectedVillageName) setColVillageNameA(detectedVillageName);
    if (detectedVillageCode) setColVillageCodeA(detectedVillageCode);
    if (detectedHouseholds) setColHouseholdsA(detectedHouseholds);
    if (detectedCommuneName) setColCommuneNameA(detectedCommuneName);
    if (detectedCommuneCode) setColCommuneCodeA(detectedCommuneCode);
  };

  const autoDetectColsB = (cols: string[]) => {
    let detectedVillageName = "";
    let detectedVillageCode = "";
    let detectedHouseholds = "";
    let detectedCommuneName = "";
    let detectedCommuneCode = "";
    let detectedDbdtName = "";
    let detectedDbdtCode = "";
    let detectedDbdtHouseholds = "";

    for (const c of cols) {
      const lower = c.toLowerCase().trim();
      const hasMa = lower.includes("mã") || lower.includes("code") || lower.startsWith("ma_") || lower.startsWith("ma ");
      const hasTen = lower.includes("tên") || lower.includes("ten");

      if (!detectedVillageName) {
        if (hasTen && (lower.includes("thôn") || lower.includes("tổ dân phố") || lower.includes("tdp"))) {
          detectedVillageName = c;
        } else if (!hasMa && (lower.includes("thôn") || lower.includes("tổ dân phố") || lower.includes("tdp")) && !lower.includes("đbđt")) {
          detectedVillageName = c;
        }
      }

      if (!detectedVillageCode) {
        if (hasMa && (lower.includes("thôn") || lower.includes("tổ dân phố") || lower.includes("tdp")) && !lower.includes("đbđt")) {
          detectedVillageCode = c;
        }
      }

      if (!detectedHouseholds) {
        if (lower.includes("tổng số hộ của thôn") || lower.includes("số hộ của thôn") || (lower.includes("tổng số hộ") && !lower.includes("đbđt"))) {
          detectedHouseholds = c;
        }
      }

      if (!detectedCommuneName) {
        if ((hasTen && (lower.includes("xã") || lower.includes("phường"))) || (!hasMa && (lower === "xã" || lower === "phường" || lower === "tên xã" || lower === "xã/phường" || lower.includes("xã / phường")))) {
          detectedCommuneName = c;
        }
      }

      if (!detectedCommuneCode) {
        if (hasMa && (lower.includes("xã") || lower.includes("phường") || lower.includes("maxa"))) {
          detectedCommuneCode = c;
        }
      }

      if (!detectedDbdtName) {
        if (lower.includes("tên đbđt") || lower.includes("đbđt") || lower.includes("địa bàn điều tra")) {
          detectedDbdtName = c;
        }
      }

      if (!detectedDbdtCode) {
        if (lower.includes("mã đbđt") || (hasMa && (lower.includes("đbđt") || lower.includes("địa bàn điều tra")))) {
          detectedDbdtCode = c;
        }
      }

      if (!detectedDbdtHouseholds) {
        if (lower.includes("hộ của đbđt") || lower.includes("số hộ của đbđt") || lower.includes("hộ đbđt")) {
          detectedDbdtHouseholds = c;
        }
      }
    }

    if (detectedVillageName) setColVillageNameB(detectedVillageName);
    if (detectedVillageCode) setColVillageCodeB(detectedVillageCode);
    if (detectedHouseholds) setColHouseholdsB(detectedHouseholds);
    if (detectedCommuneName) setColCommuneNameB(detectedCommuneName);
    if (detectedCommuneCode) setColCommuneCodeB(detectedCommuneCode);
    if (detectedDbdtName) setColDbdtNameB(detectedDbdtName);
    if (detectedDbdtCode) setColDbdtCodeB(detectedDbdtCode);
    if (detectedDbdtHouseholds) setColDbdtHouseholdsB(detectedDbdtHouseholds);
  };

  useEffect(() => {
    if (columnsA.length > 0) autoDetectColsA(columnsA);
  }, [columnsA]);

  useEffect(() => {
    if (columnsB.length > 0) autoDetectColsB(columnsB);
  }, [columnsB]);

  // Nạp tệp Excel A
  const handleUploadFileA = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileNameA(file.name);
    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const buffer = evt.target?.result;
        const wb = XLSX.read(buffer, { type: "binary" });
        const wsName = wb.SheetNames[0];
        const rawRows = XLSX.utils.sheet_to_json<any[]>(wb.Sheets[wsName], { header: 1, defval: "" });
        const parsed = parse2DArrayWithSmartHeader(rawRows);
        setDataA(parsed.data);
        setColumnsA(parsed.columns);
        autoDetectColsA(parsed.columns);
        setMatchingResults([]);
      } catch (err: any) {
        alert("Lỗi khi đọc file A: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Nạp tệp Excel B
  const handleUploadFileB = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileNameB(file.name);
    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const buffer = evt.target?.result;
        const wb = XLSX.read(buffer, { type: "binary" });
        const wsName = wb.SheetNames[0];
        const rawRows = XLSX.utils.sheet_to_json<any[]>(wb.Sheets[wsName], { header: 1, defval: "" });
        const parsed = parse2DArrayWithSmartHeader(rawRows);
        setDataB(parsed.data);
        setColumnsB(parsed.columns);
        autoDetectColsB(parsed.columns);
        setMatchingResults([]);
      } catch (err: any) {
        alert("Lỗi khi đọc file B: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Nạp mẫu tức thì từ 2 ảnh của người dùng
  const handleLoadSampleFromUserImages = () => {
    setDataA(SAMPLE_OLD_VILLAGES_FROM_IMAGE);
    const colsA = Object.keys(SAMPLE_OLD_VILLAGES_FROM_IMAGE[0]);
    setColumnsA(colsA);
    setFileNameA("Mau_DiaBan_Cu_HienNam_AnTao_LeLoi.xlsx");
    setColCommuneNameA("Tên xã");
    setColCommuneCodeA("Mã xã");
    setColVillageNameA("Tên thôn, Tổ dân phố");
    setColVillageCodeA("Mã Thôn, Tổ dân phố");
    setColHouseholdsA("Tổng số hộ của thôn, Tổ dân phố");

    setDataB(SAMPLE_NEW_VILLAGES_FROM_IMAGE);
    const colsB = Object.keys(SAMPLE_NEW_VILLAGES_FROM_IMAGE[0]);
    setColumnsB(colsB);
    setFileNameB("Mau_DiaBan_Moi_PhoHien_NamNay.xlsx");
    setColCommuneNameB("Tên xã");
    setColCommuneCodeB("Mã xã");
    setColVillageNameB("Tên thôn/TDP");
    setColVillageCodeB("Mã thôn/TDP");
    setColHouseholdsB("Tổng số hộ của thôn/TDP");
    setColDbdtNameB("Tên ĐBĐT");
    setColDbdtCodeB("Mã ĐBĐT");
    setColDbdtHouseholdsB("Tổng số hộ của ĐBĐT");

    setMatchingResults([]);
  };

  // Danh sách các xã duy nhất có trong Tệp A để người dùng chọn chạy theo từng xã
  const uniqueCommunesA = useMemo(() => {
    if (!colCommuneNameA || dataA.length === 0) return [];
    const map = new Map<string, number>();
    dataA.forEach(r => {
      const c = String(r[colCommuneNameA] || "").trim();
      if (c) map.set(c, (map.get(c) || 0) + 1);
    });
    return Array.from(map.entries())
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);
  }, [dataA, colCommuneNameA]);

  // Khởi chạy thuật toán khớp nối
  const handleRunCrosswalk = () => {
    if (dataA.length === 0 || dataB.length === 0) {
      alert("Vui lòng nạp đầy đủ Tệp Địa bàn cũ (A) và Tệp Địa bàn mới (B)!");
      return;
    }
    if (!colVillageNameA || !colVillageNameB) {
      alert("Vui lòng chọn cột Tên thôn/TDP ở cả 2 tệp!");
      return;
    }

    setIsProcessing(true);

    setTimeout(() => {
      try {
        const config: VillageAreaCrosswalkConfig = {
          colCommuneNameA,
          colCommuneCodeA,
          colVillageNameA,
          colVillageCodeA,
          colHouseholdsA,
          colCommuneNameB,
          colCommuneCodeB,
          colVillageNameB,
          colVillageCodeB,
          colHouseholdsB,
          colDbdtNameB,
          colDbdtCodeB,
          colDbdtHouseholdsB,
          customCommuneCrosswalk,
          maxHouseholdDiffPercent: 25,
          selectedCommuneA,
          strictCommuneScoped,
        };

        const { results, summary } = executeVillageAreaCrosswalk(dataA, dataB, config);
        setMatchingResults(results);
        setMatchingSummary(summary);
        setActiveViewTab("ACTIVE");
      } catch (err: any) {
        alert("Lỗi khi xử lý đối soát: " + err.message);
      } finally {
        setIsProcessing(false);
      }
    }, 50);
  };

  // Chọn thủ công 1 ứng viên
  const handleSelectCandidateManually = (rowIdx: number, candidate: any) => {
    const updated = [...matchingResults];
    const target = updated[rowIdx];
    if (!target) return;

    target.matchedB = candidate.recordB;
    target.confidence = 100;
    target.matchGrade = "PERFECT";
    target.matchReason = "Đã xác nhận thủ công bởi cán bộ thống kê";
    target.diffHouseholds = Math.abs(target.recordA.householdCountA - candidate.recordB.householdCountB);
    const maxH = Math.max(target.recordA.householdCountA, candidate.recordB.householdCountB, 1);
    target.diffPercentage = (target.diffHouseholds / maxH) * 100;
    target.manualOverride = true;

    setMatchingResults(updated);
    setSelectedCandidateRow(null);
  };

  // Hủy gán cho 1 dòng
  const handleUnmatchRow = (rowIdx: number) => {
    const updated = [...matchingResults];
    const target = updated[rowIdx];
    if (!target) return;

    target.matchedB = null;
    target.confidence = 0;
    target.matchGrade = "UNMATCHED";
    target.matchReason = "Đã hủy gán bởi người dùng";
    target.diffHouseholds = 0;
    target.diffPercentage = 0;
    target.manualOverride = true;

    setMatchingResults(updated);
  };

  // ==========================================
  // LOGIC BÓC TÁCH & KHO LƯU TRỮ (CẤT ĐI & KHÔI PHỤC)
  // "Lấy đúng ra trước cất đi, chỉ còn lại cái chưa khớp thì làm dần sẽ nhanh hơn"
  // ==========================================

  // Cất tất cả các địa bàn đã khớp đúng vào Kho Lưu Trữ
  const handleStashAllMatched = (onlyHighConfidence: boolean = false) => {
    const toStash: VillageMatchResult[] = [];
    const remaining: VillageMatchResult[] = [];

    matchingResults.forEach(r => {
      const isEligible = onlyHighConfidence
        ? r.matchGrade === "PERFECT" || r.matchGrade === "VERY_HIGH"
        : r.matchedB !== null && r.matchGrade !== "UNMATCHED";

      if (isEligible) {
        toStash.push(r);
      } else {
        remaining.push(r);
      }
    });

    if (toStash.length === 0) {
      alert("Không có địa bàn nào đã khớp trong bảng hiện tại để cất đi!");
      return;
    }

    setStashedMatches(prev => [...prev, ...toStash]);
    setMatchingResults(remaining);
    setActiveViewTab("ACTIVE");
    setAppliedMsg(`📦 Đã cất ${toStash.length} địa bàn khớp đúng vào Kho Lưu Trữ an toàn! Bảng làm việc hiện chỉ còn ${remaining.length} địa bàn chưa khớp để bạn lọc làm tiếp.`);
  };

  // Cất 1 dòng cụ thể vào kho
  const handleStashSingleRow = (indexInActive: number) => {
    const item = matchingResults[indexInActive];
    if (!item) return;
    setStashedMatches(prev => [item, ...prev]);
    setMatchingResults(prev => prev.filter((_, idx) => idx !== indexInActive));
  };

  // Khôi phục 1 dòng từ Kho Lưu Trữ về lại bảng làm việc
  const handleUnstashSingleRow = (indexInStashed: number) => {
    const item = stashedMatches[indexInStashed];
    if (!item) return;
    setMatchingResults(prev => [item, ...prev]);
    setStashedMatches(prev => prev.filter((_, idx) => idx !== indexInStashed));
  };

  // Khôi phục toàn bộ kho lưu trữ về bảng làm việc
  const handleUnstashAll = () => {
    if (stashedMatches.length === 0) return;
    setMatchingResults(prev => [...stashedMatches, ...prev]);
    setStashedMatches([]);
    setAppliedMsg("↩️ Đã khôi phục toàn bộ các địa bàn từ Kho Lưu Trữ về lại bảng làm việc!");
  };

  // Gán kết quả vào dữ liệu chính đang làm việc (HỢP NHẤT CẢ KHO LƯU TRỮ VÀ BẢNG ĐANG LÀM)
  const handleApplyCodesToWorkingTable = () => {
    const allMatches = [...stashedMatches, ...matchingResults];
    if (allMatches.length === 0 || !onApplyMergedToMain) return;

    const codeMap = new Map<number, any>();
    allMatches.forEach(r => {
      if (r.matchedB) {
        codeMap.set(r.recordA.indexA, {
          Ma_Xa_Moi: r.matchedB.communeCodeB,
          Ten_Xa_Moi: r.matchedB.communeNameB,
          Ma_Thon_Moi: r.matchedB.villageCodeB,
          Ten_Thon_Moi: r.matchedB.villageNameB,
          Tong_So_Ho_Thon_Moi: r.matchedB.householdCountB,
          Ma_DBDT_Moi: r.matchedB.dbdtCodeB,
          Ten_DBDT_Moi: r.matchedB.dbdtNameB,
          Do_Tin_Cay_DoiSoat: `${r.confidence}%`,
        });
      }
    });

    const updatedData = dataA.map((r, idx) => {
      const match = codeMap.get(idx);
      if (match) {
        return { ...r, ...match };
      }

      // Ngay cả khi chưa khớp địa bàn cụ thể, tự động gán Mã Xã Mới và Tên Xã Mới chuẩn xác từ Từ Điển 104 xã
      const commCodeA = colCommuneCodeA ? String(r[colCommuneCodeA] || "").trim() : "";
      const commNameA = colCommuneNameA ? String(r[colCommuneNameA] || "").trim() : "";
      const fallbackNewCode =
        (commCodeA && findOfficialNewCommuneCode(commCodeA)) ||
        (commNameA && findOfficialNewCommuneCode(commNameA)) ||
        "";
      const fallbackNewName =
        (commCodeA && findOfficialNewCommune(commCodeA)) ||
        (commNameA && findOfficialNewCommune(commNameA)) ||
        "";

      return {
        ...r,
        Ma_Xa_Moi: fallbackNewCode,
        Ten_Xa_Moi: fallbackNewName,
        Ma_Thon_Moi: "",
        Ten_Thon_Moi: "",
        Tong_So_Ho_Thon_Moi: "",
        Ma_DBDT_Moi: "",
        Ten_DBDT_Moi: "",
        Do_Tin_Cay_DoiSoat: fallbackNewCode ? "Khớp mã xã (thôn chưa đối soát)" : "Chưa khớp",
      };
    });

    const newCols = [...columnsA];
    [
      "Ma_Xa_Moi",
      "Ten_Xa_Moi",
      "Ma_Thon_Moi",
      "Ten_Thon_Moi",
      "Tong_So_Ho_Thon_Moi",
      "Ma_DBDT_Moi",
      "Ten_DBDT_Moi",
      "Do_Tin_Cay_DoiSoat",
    ].forEach(c => {
      if (!newCols.includes(c)) newCols.push(c);
    });

    onApplyMergedToMain(updatedData, newCols);
    setAppliedMsg(`Đã gán thành công mã xã & thôn mới cho ${codeMap.size}/${dataA.length} địa bàn vào bảng làm việc (bao gồm ${stashedMatches.length} địa bàn trong Kho Lưu Trữ)!`);
  };

  // Danh sách hiển thị theo bộ lọc và tab (Bảng đang làm hoặc Kho đã cất)
  const currentWorkingList = activeViewTab === "ACTIVE" ? matchingResults : stashedMatches;

  const filteredResults = useMemo(() => {
    return currentWorkingList.filter(r => {
      if (filterTab === "PERFECT" && r.matchGrade !== "PERFECT") return false;
      if (filterTab === "VERY_HIGH" && r.matchGrade !== "VERY_HIGH") return false;
      if (filterTab === "HIGH" && r.matchGrade !== "HIGH") return false;
      if (filterTab === "MEDIUM" && r.matchGrade !== "MEDIUM") return false;
      if (filterTab === "UNMATCHED" && r.matchGrade !== "UNMATCHED") return false;

      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase().trim();
        const vA = r.recordA.villageNameA.toLowerCase();
        const cA = r.recordA.communeNameA.toLowerCase();
        const vB = r.matchedB ? r.matchedB.villageNameB.toLowerCase() : "";
        const cB = r.matchedB ? r.matchedB.communeNameB.toLowerCase() : "";
        const dbdt = r.matchedB ? r.matchedB.dbdtNameB.toLowerCase() : "";
        return vA.includes(q) || cA.includes(q) || vB.includes(q) || cB.includes(q) || dbdt.includes(q);
      }

      return true;
    });
  }, [currentWorkingList, filterTab, searchTerm]);

  // Lọc danh mục 310 xã chính thức
  const filteredOfficialCommunes = useMemo(() => {
    if (!officialSearchQuery.trim()) return OFFICIAL_COMMUNE_MERGE_DATA;
    const q = officialSearchQuery.toLowerCase().trim();
    return OFFICIAL_COMMUNE_MERGE_DATA.filter(
      r => r.communeA.toLowerCase().includes(q) || r.communeB.toLowerCase().includes(q)
    );
  }, [officialSearchQuery]);

  // Xuất file Excel đầy đủ (Gồm cả Kho đã cất và Bảng đang làm)
  const handleExportFullExcel = () => {
    const allRecords = [...stashedMatches, ...matchingResults];
    if (allRecords.length === 0) {
      alert("Chưa có kết quả để xuất!");
      return;
    }

    const exportRows = allRecords.map(r => {
      const orig = { ...r.recordA.raw };
      return {
        ...orig,
        "--- KẾT QUẢ ĐỐI SOÁT ĐỊA BÀN MỚI ---": "➔",
        "Mã xã mới": r.matchedB?.communeCodeB || "",
        "Tên xã mới": r.matchedB?.communeNameB || "",
        "Mã thôn/TDP mới": r.matchedB?.villageCodeB || "",
        "Tên thôn/TDP mới": r.matchedB?.villageNameB || "",
        "Tổng số hộ thôn mới": r.matchedB?.householdCountB ?? "",
        "Số hộ thôn cũ": r.recordA.householdCountA,
        "Chênh lệch số hộ": r.matchedB ? r.matchedB.householdCountB - r.recordA.householdCountA : "",
        "% Lệch số hộ": r.matchedB ? `${r.diffPercentage.toFixed(1)}%` : "",
        "Mã ĐBĐT mới": r.matchedB?.dbdtCodeB || "",
        "Tên ĐBĐT mới": r.matchedB?.dbdtNameB || "",
        "Số hộ ĐBĐT mới": r.matchedB?.dbdtHouseholdCountB ?? "",
        "Độ tin cậy (%)": r.matchedB ? `${r.confidence}%` : "0%",
        "Trạng thái đối soát":
          r.matchGrade === "PERFECT"
            ? "Khớp hoàn hảo (100%)"
            : r.matchGrade === "VERY_HIGH"
            ? "Khớp rất cao (≥95%)"
            : r.matchGrade === "HIGH"
            ? "Khớp cao (≥85%)"
            : r.matchGrade === "MEDIUM"
            ? "Nghi ngờ biến động lớn"
            : "Chưa tìm thấy",
        "Lưu trữ": stashedMatches.includes(r) ? "Đã cất vào kho" : "Bảng đang làm",
        "Lý do đối soát": r.matchReason,
      };
    });

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DiaBan_DaDoiSoat_Moi");
    XLSX.writeFile(wb, `DanhSach_ToanBo_${allRecords.length}DiaBan_DaDoiSoat.xlsx`);
  };

  // Xuất file Excel chỉ riêng Kho Đã Cất Hoàn Tất
  const handleExportStashedExcel = () => {
    if (stashedMatches.length === 0) {
      alert("Kho Lưu Trữ hiện chưa có địa bàn nào được cất!");
      return;
    }

    const exportRows = stashedMatches.map(r => ({
      ...r.recordA.raw,
      "--- ĐỊA BÀN MỚI ĐÃ KHỚP ĐÚNG ---": "➔",
      "Mã xã mới": r.matchedB?.communeCodeB || "",
      "Tên xã mới": r.matchedB?.communeNameB || "",
      "Mã thôn/TDP mới": r.matchedB?.villageCodeB || "",
      "Tên thôn/TDP mới": r.matchedB?.villageNameB || "",
      "Số hộ mới": r.matchedB?.householdCountB ?? "",
      "Số hộ cũ": r.recordA.householdCountA,
      "Chênh lệch số hộ": r.matchedB ? r.matchedB.householdCountB - r.recordA.householdCountA : "",
      "Mã ĐBĐT mới": r.matchedB?.dbdtCodeB || "",
      "Tên ĐBĐT mới": r.matchedB?.dbdtNameB || "",
      "Độ tin cậy": `${r.confidence}%`,
      "Lý do": r.matchReason,
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Kho_DaCat_HoanThanh");
    XLSX.writeFile(wb, `Kho_DaCat_HoanThanh_${stashedMatches.length}DiaBan.xlsx`);
  };

  // Xuất file Excel chỉ riêng những địa bàn Còn Chưa Khớp
  const handleExportUnmatchedExcel = () => {
    const unmatched = matchingResults.filter(r => !r.matchedB || r.matchGrade === "UNMATCHED");
    if (unmatched.length === 0) {
      alert("Không có địa bàn nào chưa khớp!");
      return;
    }

    const exportRows = unmatched.map((r, i) => ({
      STT: i + 1,
      "Tên xã cũ": r.recordA.communeNameA,
      "Mã xã cũ": r.recordA.communeCodeA,
      "Tên thôn/TDP cũ": r.recordA.villageNameA,
      "Mã thôn/TDP cũ": r.recordA.villageCodeA,
      "Số hộ thôn cũ": r.recordA.householdCountA,
      "Ghi chú kiểm tra": "Chưa tìm thấy địa bàn mới - Cần đi thực địa hoặc đối soát cụm hộ",
      ...r.recordA.raw,
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DiaBan_ChuaKhop");
    XLSX.writeFile(wb, `DanhSach_${unmatched.length}DiaBan_ChuaKhop.xlsx`);
  };

  // Xuất bảng ánh xạ tra cứu (Crosswalk Lookup Sheet)
  const handleExportLookupTable = () => {
    if (matchingResults.length === 0) {
      alert("Chưa có kết quả để xuất!");
      return;
    }

    const lookupRows = matchingResults
      .filter(r => r.matchedB !== null)
      .map(r => ({
        "Mã xã cũ": r.recordA.communeCodeA,
        "Tên xã cũ": r.recordA.communeNameA,
        "Mã thôn/TDP cũ": r.recordA.villageCodeA,
        "Tên thôn/TDP cũ": r.recordA.villageNameA,
        "Số hộ cũ": r.recordA.householdCountA,
        "➔": "➔",
        "Mã xã mới": r.matchedB?.communeCodeB,
        "Tên xã mới": r.matchedB?.communeNameB,
        "Mã thôn/TDP mới": r.matchedB?.villageCodeB,
        "Tên thôn/TDP mới": r.matchedB?.villageNameB,
        "Số hộ mới": r.matchedB?.householdCountB,
        "Chênh lệch": r.matchedB ? r.matchedB.householdCountB - r.recordA.householdCountA : "",
        "Mã ĐBĐT mới": r.matchedB?.dbdtCodeB,
        "Tên ĐBĐT mới": r.matchedB?.dbdtNameB,
        "Số hộ ĐBĐT mới": r.matchedB?.dbdtHouseholdCountB,
        "Độ tin cậy": `${r.confidence}%`,
      }));

    const ws = XLSX.utils.json_to_sheet(lookupRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Bang_Anh_Xa_DiaBan");
    XLSX.writeFile(wb, `Bang_TraCuu_MaDiaBan_Cu_Sang_Moi.xlsx`);
  };

  return (
    <div className="bg-white border border-slate-300 shadow-sm space-y-4 font-sans text-slate-800">
      {/* 1. HEADER CHÍNH CỦA CÔNG CỤ */}
      <div className="bg-[#286e42] text-white p-2.5 sm:p-3 flex flex-col md:flex-row md:items-center justify-between gap-2.5">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 bg-white/10 flex items-center justify-center border border-white/20 shrink-0">
            <MapPin className="w-5 h-5 text-amber-300" />
          </div>
          <div>
            <h3 className="text-sm font-bold uppercase tracking-wider text-white m-0">
              ĐỐI SOÁT ĐỊA BÀN &amp; THÔN SÁP NHẬP
            </h3>
            <p className="text-[11px] text-emerald-100 font-normal m-0 pt-0.5">
              Tự động đối chiếu thôn và số hộ theo danh mục 310 xã sáp nhập.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-1.5 shrink-0">
          <button
            type="button"
            onClick={() => setShowCommuneClusterMatrix(true)}
            className="bg-amber-400 hover:bg-amber-500 text-slate-950 font-bold text-xs px-2.5 py-1.5 flex items-center gap-1 cursor-pointer border-0 shadow-sm"
            title="Gom xã mới: Tìm đủ toàn bộ xã cũ tương ứng để gom đủ số địa bàn"
          >
            <Layers className="w-4 h-4 text-emerald-950" />
            <span>Gom xã mới</span>
          </button>

          <button
            type="button"
            onClick={() => setShowClusterResolver(true)}
            className="bg-emerald-800 hover:bg-emerald-700 text-white font-bold text-xs px-2.5 py-1.5 flex items-center gap-1 cursor-pointer border border-emerald-400/40 shadow-xs"
            title="So khớp danh sách hộ dân giữa 2 địa bàn"
          >
            <Users className="w-4 h-4 text-amber-300" />
            <span>So cụm hộ</span>
          </button>

          <button
            type="button"
            onClick={() => setShowOfficialModal(true)}
            className="bg-emerald-900/80 hover:bg-emerald-900 text-white font-bold text-xs px-2.5 py-1.5 flex items-center gap-1 cursor-pointer border border-emerald-400/30 shadow-xs"
            title="Xem danh mục 310 xã sáp nhập"
          >
            <BookOpen className="w-4 h-4 text-emerald-300" />
            <span>Danh mục xã</span>
          </button>

          <button
            type="button"
            onClick={handleLoadSampleFromUserImages}
            className="bg-amber-400 hover:bg-amber-500 text-slate-950 font-bold text-xs px-2.5 py-1.5 flex items-center gap-1 cursor-pointer shadow-xs border-0"
            title="Nạp dữ liệu mẫu thử nghiệm"
          >
            <Sparkles className="w-3.5 h-3.5 text-slate-950" />
            <span>Dữ liệu mẫu</span>
          </button>

          {onClose && (
            <button
              type="button"
              onClick={onClose}
              className="text-white/80 hover:text-white cursor-pointer bg-white/10 hover:bg-white/20 px-2 py-1.5 text-xs border border-white/20"
            >
              ✕ Đóng
            </button>
          )}
        </div>
      </div>

      {/* THÔNG BÁO GÁN MÃ VÀO BẢNG LÀM VIỆC */}
      {appliedMsg && (
        <div className="mx-3.5 p-2.5 bg-emerald-50 border-2 border-emerald-500 text-emerald-950 text-xs font-bold flex items-center justify-between gap-2 shadow-xs">
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

      {/* 2. KHUNG CHỌN TỆP & CẤU HÌNH CỘT (2 CỘT SONG SONG) */}
      <div className="p-3 sm:p-4 space-y-3">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
          {/* TỆP A: ĐỊA BÀN CŨ (LÚC CHƯA SÁP NHẬP) */}
          <div className="bg-sky-50/50 border border-sky-300 p-2.5 space-y-2.5">
            <div className="flex items-center justify-between border-b border-sky-200 pb-1.5">
              <div>
                <span className="text-xs font-bold text-sky-950 uppercase flex items-center gap-1.5">
                  <FileSpreadsheet className="w-4 h-4 text-sky-700" />
                  1. Tệp A: Địa bàn cũ
                </span>
                <span className="text-[11px] text-sky-800 font-medium block">
                  {dataA.length > 0 ? `${dataA.length.toLocaleString("vi-VN")} dòng (${fileNameA})` : "Chưa có dữ liệu"}
                </span>
              </div>

              <div>
                <input
                  type="file"
                  ref={fileInputRefA}
                  accept=".xlsx,.xls,.csv"
                  onChange={handleUploadFileA}
                  className="hidden"
                />
                <button
                  type="button"
                  onClick={() => fileInputRefA.current?.click()}
                  className="bg-sky-700 hover:bg-sky-800 text-white font-bold text-xs px-2.5 py-1 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                >
                  <Upload className="w-3.5 h-3.5" />
                  {dataA.length > 0 ? "Đổi tệp A" : "Tải tệp A"}
                </button>
              </div>
            </div>

            {columnsA.length > 0 ? (
              <div className="space-y-2 text-xs">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <div>
                    <label className="text-[11px] font-bold text-slate-700 block">
                      Tên thôn/TDP <span className="text-rose-600">*</span>:
                    </label>
                    <select
                      value={colVillageNameA}
                      onChange={e => setColVillageNameA(e.target.value)}
                      className="w-full border border-sky-300 bg-white p-1 text-xs font-bold text-sky-900"
                    >
                      <option value="">-- Chọn cột --</option>
                      {columnsA.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="text-[11px] font-bold text-emerald-800 block">
                      Tổng số hộ <span className="text-rose-600">*</span>:
                    </label>
                    <select
                      value={colHouseholdsA}
                      onChange={e => setColHouseholdsA(e.target.value)}
                      className="w-full border border-emerald-400 bg-emerald-50/50 p-1 text-xs font-bold text-emerald-950"
                    >
                      <option value="">-- Chọn cột số hộ --</option>
                      {columnsA.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 pt-0.5">
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Tên xã cũ:</label>
                    <select
                      value={colCommuneNameA}
                      onChange={e => setColCommuneNameA(e.target.value)}
                      className="w-full border border-slate-300 bg-white p-1 text-[11px]"
                    >
                      <option value="">(Không chọn)</option>
                      {columnsA.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Mã xã cũ:</label>
                    <select
                      value={colCommuneCodeA}
                      onChange={e => setColCommuneCodeA(e.target.value)}
                      className="w-full border border-slate-300 bg-white p-1 text-[11px]"
                    >
                      <option value="">(Không chọn)</option>
                      {columnsA.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Mã thôn cũ:</label>
                    <select
                      value={colVillageCodeA}
                      onChange={e => setColVillageCodeA(e.target.value)}
                      className="w-full border border-slate-300 bg-white p-1 text-[11px]"
                    >
                      <option value="">(Không chọn)</option>
                      {columnsA.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            ) : (
              <div className="text-[11px] text-slate-500 italic p-3 text-center bg-white border border-dashed border-sky-200">
                Hãy tải lên tệp địa bàn cũ hoặc bấm &ldquo;Dữ liệu mẫu&rdquo; ở trên.
              </div>
            )}
          </div>

          {/* TỆP B: ĐỊA BÀN MỚI NĂM NAY (ĐÃ SÁP NHẬP) */}
          <div className="bg-emerald-50/50 border border-emerald-300 p-2.5 space-y-2.5">
            <div className="flex items-center justify-between border-b border-emerald-200 pb-1.5">
              <div>
                <span className="text-xs font-bold text-emerald-950 uppercase flex items-center gap-1.5">
                  <FileSpreadsheet className="w-4 h-4 text-emerald-700" />
                  2. Tệp B: Địa bàn mới
                </span>
                <span className="text-[11px] text-emerald-800 font-medium block">
                  {dataB.length > 0 ? `${dataB.length.toLocaleString("vi-VN")} dòng (${fileNameB})` : "Chưa có dữ liệu"}
                </span>
              </div>

              <div>
                <input
                  type="file"
                  ref={fileInputRefB}
                  accept=".xlsx,.xls,.csv"
                  onChange={handleUploadFileB}
                  className="hidden"
                />
                <button
                  type="button"
                  onClick={() => fileInputRefB.current?.click()}
                  className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs px-2.5 py-1 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                >
                  <Upload className="w-3.5 h-3.5" />
                  {dataB.length > 0 ? "Đổi tệp B" : "Tải tệp B"}
                </button>
              </div>
            </div>

            {columnsB.length > 0 ? (
              <div className="space-y-2 text-xs">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <div>
                    <label className="text-[11px] font-bold text-slate-700 block">
                      Tên thôn/TDP mới <span className="text-rose-600">*</span>:
                    </label>
                    <select
                      value={colVillageNameB}
                      onChange={e => setColVillageNameB(e.target.value)}
                      className="w-full border border-emerald-300 bg-white p-1 text-xs font-bold text-emerald-900"
                    >
                      <option value="">-- Chọn cột --</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="text-[11px] font-bold text-emerald-800 block">
                      Tổng số hộ mới <span className="text-rose-600">*</span>:
                    </label>
                    <select
                      value={colHouseholdsB}
                      onChange={e => setColHouseholdsB(e.target.value)}
                      className="w-full border border-emerald-400 bg-emerald-50/50 p-1 text-xs font-bold text-emerald-950"
                    >
                      <option value="">-- Chọn cột số hộ --</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 pt-0.5">
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Tên xã mới:</label>
                    <select
                      value={colCommuneNameB}
                      onChange={e => setColCommuneNameB(e.target.value)}
                      className="w-full border border-slate-300 bg-white p-1 text-[11px]"
                    >
                      <option value="">(Không chọn)</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Mã xã mới:</label>
                    <select
                      value={colCommuneCodeB}
                      onChange={e => setColCommuneCodeB(e.target.value)}
                      className="w-full border border-slate-300 bg-white p-1 text-[11px]"
                    >
                      <option value="">(Không chọn)</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Mã thôn mới:</label>
                    <select
                      value={colVillageCodeB}
                      onChange={e => setColVillageCodeB(e.target.value)}
                      className="w-full border border-slate-300 bg-white p-1 text-[11px]"
                    >
                      <option value="">(Không chọn)</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Các cột ĐBĐT (Địa bàn điều tra) */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 pt-1 border-t border-emerald-200">
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Tên ĐB:</label>
                    <select
                      value={colDbdtNameB}
                      onChange={e => setColDbdtNameB(e.target.value)}
                      className="w-full border border-slate-300 bg-white p-1 text-[11px]"
                    >
                      <option value="">(Không chọn)</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Mã ĐB:</label>
                    <select
                      value={colDbdtCodeB}
                      onChange={e => setColDbdtCodeB(e.target.value)}
                      className="w-full border border-slate-300 bg-white p-1 text-[11px]"
                    >
                      <option value="">(Không chọn)</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Số hộ ĐB:</label>
                    <select
                      value={colDbdtHouseholdsB}
                      onChange={e => setColDbdtHouseholdsB(e.target.value)}
                      className="w-full border border-slate-300 bg-white p-1 text-[11px]"
                    >
                      <option value="">(Không chọn)</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            ) : (
              <div className="text-[11px] text-slate-500 italic p-3 text-center bg-white border border-dashed border-emerald-200">
                Hãy tải lên tệp địa bàn mới hoặc bấm &ldquo;Dữ liệu mẫu&rdquo; ở trên.
              </div>
            )}
          </div>
        </div>

        {/* KHUNG CẤU HÌNH PHẠM VI THEO XÃ (CHẠY THEO XÃ & NỘI BỘ XÃ VỚI NHAU) */}
        <div className="bg-amber-50/80 border border-amber-300 p-2.5 space-y-2">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-1.5">
              <Building2 className="w-4 h-4 text-amber-800" />
              <span className="text-xs font-bold text-amber-950 uppercase tracking-wide">
                Lọc theo xã:
              </span>
            </div>

            <label className="flex items-center gap-1.5 text-xs text-amber-950 font-bold cursor-pointer select-none bg-white px-2 py-0.5 border border-amber-300">
              <input
                type="checkbox"
                checked={strictCommuneScoped}
                onChange={e => setStrictCommuneScoped(e.target.checked)}
                className="w-3.5 h-3.5 text-emerald-700 rounded cursor-pointer"
              />
              <span>🔒 Khóa theo xã sáp nhập</span>
            </label>
          </div>

          <div className="flex flex-wrap items-center gap-2 text-xs">
            <div className="flex items-center gap-1.5 flex-1 min-w-[260px]">
              <label className="font-bold text-slate-700 shrink-0 text-xs">Xã:</label>
              <select
                value={selectedCommuneA}
                onChange={e => setSelectedCommuneA(e.target.value)}
                className="w-full bg-white border border-amber-400 p-1 text-xs font-bold text-slate-900 shadow-2xs"
              >
                <option value="ALL">Tất cả các xã ({uniqueCommunesA.length} xã, {dataA.length.toLocaleString("vi-VN")} địa bàn)</option>
                {uniqueCommunesA.map(c => (
                  <option key={c.name} value={c.name}>
                    📍 {c.name} ({c.count} địa bàn)
                  </option>
                ))}
              </select>
            </div>

            {selectedCommuneA !== "ALL" && (
              <button
                type="button"
                onClick={() => setSelectedCommuneA("ALL")}
                className="bg-slate-200 hover:bg-slate-300 text-slate-800 px-2.5 py-1 text-xs font-bold cursor-pointer border border-slate-300"
              >
                ✕ Tất cả
              </button>
            )}
          </div>
        </div>

        {/* NÚT KHỞI CHẠY ĐỐI SOÁT */}
        <div className="flex flex-wrap items-center justify-between gap-2 bg-slate-100 p-2.5 border border-slate-300">
          <div className="text-xs text-slate-600">
            Tự động so khớp theo tên thôn &amp; số hộ theo xã sáp nhập.
          </div>

          <button
            type="button"
            onClick={handleRunCrosswalk}
            disabled={isProcessing || dataA.length === 0 || dataB.length === 0}
            className="bg-[#286e42] hover:bg-[#205835] disabled:bg-slate-300 text-white font-bold text-xs px-5 py-2 flex items-center gap-1.5 cursor-pointer shadow-sm border-0 transition-all active:scale-95"
          >
            {isProcessing ? (
              <>
                <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                <span>Đang rà soát...</span>
              </>
            ) : (
              <>
                <GitMerge className="w-3.5 h-3.5" />
                <span>TIẾN HÀNH ĐỐI SOÁT</span>
              </>
            )}
          </button>
        </div>

        {/* 3. KẾT QUẢ TỔNG HỢP & BẢNG ĐỐI SOÁT CHI TIẾT */}
        {matchingSummary && (
          <div className="space-y-4 pt-2">
            {/* THẺ TỔNG HỢP CHỈ SỐ */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2">
              <div
                onClick={() => setFilterTab("ALL")}
                className={`p-2.5 border cursor-pointer transition-all ${
                  filterTab === "ALL" ? "border-slate-800 bg-slate-800 text-white shadow-xs" : "border-slate-300 bg-slate-50 hover:bg-slate-100"
                }`}
              >
                <div className="text-[10.5px] uppercase font-bold opacity-80">Tổng Thôn Cũ</div>
                <div className="text-lg font-black">{matchingSummary.totalA}</div>
                <div className="text-[10px] opacity-70">Tỷ lệ khớp: {matchingSummary.matchRate}%</div>
              </div>

              <div
                onClick={() => setFilterTab("PERFECT")}
                className={`p-2.5 border cursor-pointer transition-all ${
                  filterTab === "PERFECT" ? "border-emerald-700 bg-emerald-700 text-white shadow-xs" : "border-emerald-300 bg-emerald-50 hover:bg-emerald-100/70"
                }`}
              >
                <div className="text-[10.5px] uppercase font-bold text-emerald-950">Khớp Tuyệt Đối (100%)</div>
                <div className="text-lg font-black text-emerald-950">{matchingSummary.perfectMatches}</div>
                <div className="text-[10px] text-emerald-800">Trùng tên + Bằng số hộ</div>
              </div>

              <div
                onClick={() => setFilterTab("VERY_HIGH")}
                className={`p-2.5 border cursor-pointer transition-all ${
                  filterTab === "VERY_HIGH" ? "border-sky-700 bg-sky-700 text-white shadow-xs" : "border-sky-300 bg-sky-50 hover:bg-sky-100/70"
                }`}
              >
                <div className="text-[10.5px] uppercase font-bold text-sky-950">Khớp Rất Cao (≥95%)</div>
                <div className="text-lg font-black text-sky-950">{matchingSummary.veryHighMatches}</div>
                <div className="text-[10px] text-sky-800">Số hộ lệch ≤ 5% (1-5 hộ)</div>
              </div>

              <div
                onClick={() => setFilterTab("HIGH")}
                className={`p-2.5 border cursor-pointer transition-all ${
                  filterTab === "HIGH" ? "border-indigo-700 bg-indigo-700 text-white shadow-xs" : "border-indigo-300 bg-indigo-50 hover:bg-indigo-100/70"
                }`}
              >
                <div className="text-[10.5px] uppercase font-bold text-indigo-950">Khớp Cao (≥85%)</div>
                <div className="text-lg font-black text-indigo-950">{matchingSummary.highMatches}</div>
                <div className="text-[10px] text-indigo-800">Số hộ lệch ≤ 15%</div>
              </div>

              <div
                onClick={() => setFilterTab("MEDIUM")}
                className={`p-2.5 border cursor-pointer transition-all ${
                  filterTab === "MEDIUM" ? "border-amber-700 bg-amber-700 text-white shadow-xs" : "border-amber-300 bg-amber-50 hover:bg-amber-100/70"
                }`}
              >
                <div className="text-[10.5px] uppercase font-bold text-amber-950">Cần Kiểm Tra</div>
                <div className="text-lg font-black text-amber-950">{matchingSummary.mediumMatches}</div>
                <div className="text-[10px] text-amber-800">Biến động số hộ lớn</div>
              </div>

              <div
                onClick={() => setFilterTab("UNMATCHED")}
                className={`p-2.5 border cursor-pointer transition-all ${
                  filterTab === "UNMATCHED" ? "border-rose-700 bg-rose-700 text-white shadow-xs" : "border-rose-300 bg-rose-50 hover:bg-rose-100/70"
                }`}
              >
                <div className="text-[10.5px] uppercase font-bold text-rose-950">Chưa Tìm Thấy</div>
                <div className="text-lg font-black text-rose-950">{matchingSummary.unmatched}</div>
                <div className="text-[10px] text-rose-800">Cần gắn tay</div>
              </div>
            </div>

            {/* KHUNG BÓC TÁCH & KHO LƯU TRỮ */}
            <div className="bg-emerald-50/70 border border-emerald-300 p-2.5 space-y-2 shadow-xs">
              <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-2">
                <div>
                  <div className="flex items-center gap-1.5">
                    <Archive className="w-4 h-4 text-emerald-800" />
                    <span className="text-xs font-bold text-emerald-950 uppercase tracking-wide">
                      Kho lưu trữ dòng đã khớp
                    </span>
                  </div>
                  <p className="text-[11px] text-emerald-800 m-0 pt-0.5">
                    Cất các dòng đã khớp đúng để bảng làm việc thu gọn lại.
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                  {matchingResults.filter(r => r.matchedB !== null && r.matchGrade !== "UNMATCHED").length > 0 && (
                    <button
                      type="button"
                      onClick={() => handleStashAllMatched(false)}
                      className="bg-emerald-800 hover:bg-emerald-900 text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer shadow-xs border border-emerald-600 active:scale-95"
                      title="Cất tất cả các địa bàn đã khớp vào kho"
                    >
                      <Archive className="w-3.5 h-3.5 text-amber-300" />
                      <span>Cất đã khớp ({matchingResults.filter(r => r.matchedB !== null && r.matchGrade !== "UNMATCHED").length})</span>
                    </button>
                  )}

                  {matchingResults.filter(r => r.matchGrade === "PERFECT" || r.matchGrade === "VERY_HIGH").length > 0 && (
                    <button
                      type="button"
                      onClick={() => handleStashAllMatched(true)}
                      className="bg-sky-800 hover:bg-sky-900 text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer shadow-xs border border-sky-600 active:scale-95"
                      title="Chỉ cất các địa bàn khớp ≥95%"
                    >
                      <CheckSquare className="w-3.5 h-3.5 text-sky-200" />
                      <span>Cất khớp ≥95% ({matchingResults.filter(r => r.matchGrade === "PERFECT" || r.matchGrade === "VERY_HIGH").length})</span>
                    </button>
                  )}
                </div>
              </div>

              {/* TAB CHUYỂN ĐỔI GIỮA BẢNG ĐANG LÀM VÀ KHO LƯU TRỮ */}
              <div className="flex flex-wrap items-center justify-between gap-2 border-t border-emerald-200 pt-2.5">
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setActiveViewTab("ACTIVE")}
                    className={`text-xs font-bold px-3.5 py-1.5 flex items-center gap-2 cursor-pointer border transition-all ${
                      activeViewTab === "ACTIVE"
                        ? "bg-slate-900 text-white border-slate-900 shadow-xs"
                        : "bg-white text-slate-700 border-slate-300 hover:bg-slate-100"
                    }`}
                  >
                    <span>⏳ Bảng Đang Làm: Còn {matchingResults.filter(r => !r.matchedB || r.matchGrade === "UNMATCHED").length} Chưa Khớp</span>
                    <span className={`text-[10.5px] px-1.5 py-0.5 rounded-xs font-mono font-bold ${activeViewTab === "ACTIVE" ? "bg-amber-400 text-slate-950" : "bg-slate-200 text-slate-800"}`}>
                      {matchingResults.length}
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => setActiveViewTab("STASHED")}
                    className={`text-xs font-bold px-3.5 py-1.5 flex items-center gap-2 cursor-pointer border transition-all ${
                      activeViewTab === "STASHED"
                        ? "bg-emerald-800 text-white border-emerald-800 shadow-xs"
                        : "bg-white text-emerald-950 border-emerald-300 hover:bg-emerald-100/50"
                    }`}
                  >
                    <FolderCheck className="w-3.5 h-3.5 text-emerald-300" />
                    <span>✅ Kho Đã Cất Hoàn Thành</span>
                    <span className={`text-[10.5px] px-1.5 py-0.5 rounded-xs font-mono font-bold ${activeViewTab === "STASHED" ? "bg-amber-300 text-slate-950" : "bg-emerald-200 text-emerald-950"}`}>
                      {stashedMatches.length}
                    </span>
                  </button>
                </div>

                {activeViewTab === "STASHED" && stashedMatches.length > 0 && (
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={handleExportStashedExcel}
                      className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs px-2.5 py-1 flex items-center gap-1 cursor-pointer border-0"
                    >
                      <Download className="w-3.5 h-3.5" />
                      <span>Xuất Excel Kho Đã Cất ({stashedMatches.length})</span>
                    </button>
                    <button
                      type="button"
                      onClick={handleUnstashAll}
                      className="bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold text-xs px-2.5 py-1 flex items-center gap-1 cursor-pointer border border-slate-300"
                    >
                      <RotateCcw className="w-3.5 h-3.5" />
                      <span>Khôi phục toàn bộ về bảng làm việc</span>
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* GỢI Ý CỨU CÁNH: ĐỐI SOÁT CỤM HỘ DÂN CHO NHỮNG THÔN CHƯA TÌM THẤY */}
            {matchingSummary.unmatched > 0 && (
              <div className="p-3 bg-amber-50 border-2 border-amber-400 flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs shadow-xs">
                <div className="flex items-start gap-2.5">
                  <div className="w-8 h-8 bg-amber-500 text-slate-950 flex items-center justify-center font-black shrink-0 shadow-2xs">
                    <Users className="w-5 h-5" />
                  </div>
                  <div>
                    <div className="font-black text-amber-950 uppercase text-xs">
                      Còn {matchingSummary.unmatched} Thôn Chưa Tìm Thấy Do Đổi Tên Hoàn Toàn Hoặc Mất Mã?
                    </div>
                    <div className="text-slate-700 pt-0.5 leading-relaxed">
                      👉 Hãy dùng giải pháp <strong>Đối Soát Cụm Hộ Dân (≥ 30 Người Trùng Tên)</strong>: So sánh danh sách hộ gia đình giữa địa bàn này với các địa bàn khác. Nếu phát hiện có từ <strong>30 người trùng tên trở lên</strong> ➔ Hệ thống tự động xác định và tạm tính là cùng một địa bàn!
                    </div>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setShowClusterResolver(true)}
                  className="bg-[#1e5430] hover:bg-[#163f24] text-white font-bold text-xs px-3.5 py-2 flex items-center gap-1.5 cursor-pointer shrink-0 shadow-xs border-0"
                >
                  <Users className="w-4 h-4 text-amber-300" />
                  <span>🚀 Mở Đối Soát Cụm Hộ Dân (≥ 30 Người Trùng)</span>
                </button>
              </div>
            )}

            {/* THANH CÔNG CỤ TÌM KIẾM & XUẤT EXCEL */}
            <div className="flex flex-wrap items-center justify-between gap-3 bg-white border border-slate-300 p-2.5">
              <div className="flex items-center gap-2 flex-1 max-w-md">
                <Search className="w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  value={searchTerm}
                  onChange={e => setSearchTerm(e.target.value)}
                  placeholder="Tìm kiếm theo tên thôn cũ, thôn mới, xã mới, ĐBĐT..."
                  className="w-full text-xs border border-slate-300 px-2 py-1 bg-white"
                />
              </div>

              <div className="flex flex-wrap items-center gap-2">
                {onApplyMergedToMain && (
                  <button
                    type="button"
                    onClick={handleApplyCodesToWorkingTable}
                    className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1.5 cursor-pointer border-0 shadow-xs"
                    title="Gán trực tiếp các cột mã xã mới, thôn mới vào bảng dữ liệu chính"
                  >
                    <Database className="w-3.5 h-3.5" />
                    <span>Gán Mã Mới Vào Dữ Liệu Kỳ 1</span>
                  </button>
                )}

                {matchingResults.some(r => !r.matchedB || r.matchGrade === "UNMATCHED") && (
                  <button
                    type="button"
                    onClick={handleExportUnmatchedExcel}
                    className="bg-rose-700 hover:bg-rose-800 text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                    title="Xuất riêng danh sách các địa bàn chưa khớp để in ra rà soát hoặc đi thực địa"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Xuất DS Chưa Khớp ({matchingResults.filter(r => !r.matchedB || r.matchGrade === "UNMATCHED").length})</span>
                  </button>
                )}

                <button
                  type="button"
                  onClick={handleExportLookupTable}
                  className="bg-slate-700 hover:bg-slate-800 text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                  title="Xuất bảng tra cứu mã: Mã thôn cũ ➔ Mã thôn mới, Mã ĐBĐT"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Xuất Bảng Tra Cứu (Crosswalk)</span>
                </button>

                <button
                  type="button"
                  onClick={handleExportFullExcel}
                  className="bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                  title="Xuất tệp Excel đầy đủ gồm toàn bộ dữ liệu (cả bảng đang làm và kho đã cất)"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Xuất Toàn Bộ Đã Gán Mã</span>
                </button>
              </div>
            </div>

            {/* BẢNG ĐỐI SOÁT ĐỊA BÀN CHI TIẾT */}
            <div className="border border-slate-300 overflow-x-auto max-h-[500px]">
              <table className="w-full text-left border-collapse text-xs">
                <thead className="bg-slate-100 text-slate-700 font-bold sticky top-0 border-b border-slate-300 z-10">
                  <tr>
                    <th className="p-2 border-r border-slate-200 text-center w-12">STT</th>
                    <th className="p-2 border-r border-slate-200 bg-sky-100/70 text-sky-950">THÔN CŨ (KỲ NN)</th>
                    <th className="p-2 border-r border-slate-200 bg-sky-100/70 text-sky-950 text-right w-24">SỐ HỘ CŨ</th>
                    <th className="p-2 border-r border-slate-200 text-center w-8">➔</th>
                    <th className={`p-2 border-r border-slate-200 ${activeViewTab === "STASHED" ? "bg-amber-100 text-amber-950" : "bg-emerald-100/70 text-emerald-950"}`}>
                      {activeViewTab === "STASHED" ? "THÔN MỚI (TRONG KHO ĐÃ CẤT)" : "THÔN MỚI (KỲ CÁ THỂ)"}
                    </th>
                    <th className={`p-2 border-r border-slate-200 text-right w-24 ${activeViewTab === "STASHED" ? "bg-amber-100 text-amber-950" : "bg-emerald-100/70 text-emerald-950"}`}>SỐ HỘ MỚI</th>
                    <th className="p-2 border-r border-slate-200 text-center w-24">CHÊNH LỆCH</th>
                    <th className="p-2 border-r border-slate-200 bg-emerald-50 text-emerald-900">ĐBĐT MỚI</th>
                    <th className="p-2 border-r border-slate-200 text-center w-28">ĐỘ TIN CẬY</th>
                    <th className="p-2 text-center w-36">THAO TÁC</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {filteredResults.length === 0 ? (
                    <tr>
                      <td colSpan={10} className="p-6 text-center text-slate-500 italic">
                        {activeViewTab === "STASHED"
                          ? "Kho Lưu Trữ hiện đang trống. Hãy bấm 'Cất vào kho' ở bảng đang làm để chuyển các địa bàn đúng vào đây."
                          : "Không tìm thấy địa bàn nào phù hợp với bộ lọc hiện tại."}
                      </td>
                    </tr>
                  ) : (
                    filteredResults.map((r, idx) => {
                      const origIndex = activeViewTab === "ACTIVE" ? matchingResults.indexOf(r) : stashedMatches.indexOf(r);
                      const isPerfect = r.matchGrade === "PERFECT";
                      const isVeryHigh = r.matchGrade === "VERY_HIGH";
                      const isHigh = r.matchGrade === "HIGH";
                      const isMedium = r.matchGrade === "MEDIUM";
                      const isUnmatched = r.matchGrade === "UNMATCHED";

                      const diffH = r.matchedB ? r.matchedB.householdCountB - r.recordA.householdCountA : 0;

                      return (
                        <tr
                          key={idx}
                          className={`hover:bg-slate-50 transition-colors ${
                            isPerfect
                              ? "bg-emerald-50/20"
                              : isVeryHigh
                              ? "bg-sky-50/20"
                              : isMedium
                              ? "bg-amber-50/30"
                              : isUnmatched
                              ? "bg-rose-50/20"
                              : ""
                          }`}
                        >
                          <td className="p-2 text-center border-r border-slate-200 text-slate-500 font-mono">
                            {idx + 1}
                          </td>

                          {/* Thôn cũ */}
                          <td className="p-2 border-r border-slate-200">
                            <div className="font-bold text-slate-900">{r.recordA.villageNameA}</div>
                            <div className="text-[11px] text-slate-500 flex items-center gap-1">
                              <span>{r.recordA.communeNameA}</span>
                              {r.recordA.villageCodeA && <span className="font-mono text-slate-400">• Mã: {r.recordA.villageCodeA}</span>}
                            </div>
                          </td>

                          {/* Số hộ cũ */}
                          <td className="p-2 border-r border-slate-200 text-right font-mono font-bold text-sky-900">
                            {r.recordA.householdCountA ? r.recordA.householdCountA.toLocaleString("vi-VN") : "0"}
                          </td>

                          <td className="p-2 border-r border-slate-200 text-center text-slate-400 font-bold">
                            ➔
                          </td>

                          {/* Thôn mới */}
                          <td className="p-2 border-r border-slate-200">
                            {r.matchedB ? (
                              <div>
                                <div className="font-bold text-emerald-950 flex items-center gap-1.5">
                                  <span>{r.matchedB.villageNameB}</span>
                                  {r.manualOverride && (
                                    <span className="text-[9px] bg-indigo-100 text-indigo-800 font-bold px-1">Gán tay</span>
                                  )}
                                </div>
                                <div className="text-[11px] text-slate-500">
                                  <span>{r.matchedB.communeNameB}</span>
                                  {r.matchedB.villageCodeB && (
                                    <span className="font-mono text-slate-400">• Mã: {r.matchedB.villageCodeB}</span>
                                  )}
                                </div>
                              </div>
                            ) : (
                              <span className="text-rose-700 italic font-semibold">Chưa tìm thấy địa bàn mới</span>
                            )}
                          </td>

                          {/* Số hộ mới */}
                          <td className="p-2 border-r border-slate-200 text-right font-mono font-bold text-emerald-900">
                            {r.matchedB && r.matchedB.householdCountB
                              ? r.matchedB.householdCountB.toLocaleString("vi-VN")
                              : "—"}
                          </td>

                          {/* Chênh lệch */}
                          <td className="p-2 border-r border-slate-200 text-center font-mono">
                            {r.matchedB ? (
                              diffH === 0 ? (
                                <span className="text-emerald-700 font-bold">0 (Bằng nhau)</span>
                              ) : (
                                <span
                                  className={Math.abs(diffH) <= 5 ? "text-sky-700 font-semibold" : "text-amber-700 font-bold"}
                                >
                                  {diffH > 0 ? `+${diffH}` : diffH} ({r.diffPercentage.toFixed(1)}%)
                                </span>
                              )
                            ) : (
                              "—"
                            )}
                          </td>

                          {/* ĐBĐT */}
                          <td className="p-2 border-r border-slate-200">
                            {r.matchedB?.dbdtNameB ? (
                              <div>
                                <span className="font-semibold text-slate-800">{r.matchedB.dbdtNameB}</span>
                                {r.matchedB.dbdtHouseholdCountB > 0 && (
                                  <span className="text-[10px] text-slate-500 block">
                                    {r.matchedB.dbdtHouseholdCountB} hộ
                                  </span>
                                )}
                              </div>
                            ) : (
                              <span className="text-slate-400 italic">Theo Thôn</span>
                            )}
                          </td>

                          {/* Độ tin cậy */}
                          <td className="p-2 border-r border-slate-200 text-center">
                            {isPerfect && (
                              <span className="bg-emerald-100 text-emerald-900 font-extrabold text-[10px] px-2 py-0.5 border border-emerald-300">
                                100% Hoàn hảo
                              </span>
                            )}
                            {isVeryHigh && (
                              <span className="bg-sky-100 text-sky-900 font-bold text-[10px] px-2 py-0.5 border border-sky-300">
                                {r.confidence}% Rất cao
                              </span>
                            )}
                            {isHigh && (
                              <span className="bg-indigo-100 text-indigo-900 font-bold text-[10px] px-2 py-0.5 border border-indigo-300">
                                {r.confidence}% Khớp cao
                              </span>
                            )}
                            {isMedium && (
                              <span className="bg-amber-100 text-amber-900 font-bold text-[10px] px-2 py-0.5 border border-amber-300">
                                {r.confidence}% Kiểm tra
                              </span>
                            )}
                            {isUnmatched && (
                              <span className="bg-rose-100 text-rose-900 font-bold text-[10px] px-2 py-0.5 border border-rose-300">
                                Chưa khớp
                              </span>
                            )}
                          </td>

                          {/* Thao tác */}
                          <td className="p-2 text-center space-x-1 whitespace-nowrap">
                            {activeViewTab === "ACTIVE" ? (
                              <>
                                {r.matchedB && (
                                  <button
                                    type="button"
                                    onClick={() => handleStashSingleRow(origIndex)}
                                    className="text-[10.5px] bg-emerald-100 hover:bg-emerald-200 text-emerald-900 font-bold px-2 py-1 border border-emerald-300 cursor-pointer inline-flex items-center gap-1"
                                    title="Cất địa bàn này vào Kho Lưu Trữ an toàn để thu gọn bảng đang làm"
                                  >
                                    <Archive className="w-3 h-3 text-emerald-800" />
                                    <span>Cất</span>
                                  </button>
                                )}

                                {r.candidates.length > 0 && (
                                  <button
                                    type="button"
                                    onClick={() => setSelectedCandidateRow({ index: origIndex, item: r })}
                                    className="text-[10.5px] bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold px-2 py-1 border border-slate-300 cursor-pointer"
                                    title="Xem danh sách các địa bàn mới ứng viên và chọn thủ công"
                                  >
                                    Ứng viên ({r.candidates.length})
                                  </button>
                                )}

                                {r.matchedB && (
                                  <button
                                    type="button"
                                    onClick={() => handleUnmatchRow(origIndex)}
                                    className="text-[10.5px] text-rose-700 hover:text-rose-900 font-bold px-1.5 py-1 cursor-pointer border-0 bg-transparent"
                                    title="Hủy gán địa bàn này"
                                  >
                                    ✕
                                  </button>
                                )}
                              </>
                            ) : (
                              <button
                                type="button"
                                onClick={() => handleUnstashSingleRow(origIndex)}
                                className="text-[10.5px] bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold px-2.5 py-1 border border-slate-300 cursor-pointer inline-flex items-center gap-1"
                                title="Khôi phục địa bàn này từ Kho Lưu Trữ về lại bảng làm việc"
                              >
                                <RotateCcw className="w-3 h-3 text-slate-600" />
                                <span>Khôi phục</span>
                              </button>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      {/* MODAL 1: CHỌN ỨNG VIÊN THỦ CÔNG KHI CẦN KIỂM TRA */}
      {selectedCandidateRow && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white border-2 border-slate-800 shadow-xl max-w-2xl w-full max-h-[85vh] flex flex-col">
            <div className="bg-slate-900 text-white p-3 flex items-center justify-between">
              <div>
                <h4 className="font-bold text-xs uppercase text-amber-300">
                  Chọn Địa Bàn Mới Cho Thôn: &ldquo;{selectedCandidateRow.item.recordA.villageNameA}&rdquo;
                </h4>
                <div className="text-[11px] text-slate-300">
                  Xã cũ: {selectedCandidateRow.item.recordA.communeNameA} • Số hộ cũ:{" "}
                  <strong>{selectedCandidateRow.item.recordA.householdCountA} hộ</strong>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setSelectedCandidateRow(null)}
                className="text-white hover:text-rose-400 font-bold text-sm bg-transparent border-0 cursor-pointer"
              >
                ✕
              </button>
            </div>

            <div className="p-3 overflow-y-auto space-y-2 flex-1">
              <div className="text-xs text-slate-600 mb-2">
                Dưới đây là các địa bàn mới có tên tương đồng hoặc cùng xã mục tiêu được sắp xếp theo độ tin cậy. Nhấn &ldquo;Chọn gán&rdquo; để xác nhận:
              </div>

              {selectedCandidateRow.item.candidates.map((cand, cIdx) => {
                const diff = cand.recordB.householdCountB - selectedCandidateRow.item.recordA.householdCountA;
                return (
                  <div
                    key={cIdx}
                    className="p-2.5 border border-slate-300 hover:border-emerald-600 hover:bg-emerald-50/40 transition-colors flex items-center justify-between gap-3"
                  >
                    <div>
                      <div className="font-bold text-slate-900 text-xs flex items-center gap-2">
                        <span>{cand.recordB.villageNameB}</span>
                        <span className="text-[10px] bg-slate-100 text-slate-700 px-1.5 py-0.5 border border-slate-300">
                          Xã mới: {cand.recordB.communeNameB}
                        </span>
                        {cand.recordB.dbdtNameB && (
                          <span className="text-[10px] bg-emerald-100 text-emerald-900 px-1.5 py-0.5">
                            ĐBĐT: {cand.recordB.dbdtNameB}
                          </span>
                        )}
                      </div>
                      <div className="text-[11px] text-slate-600 mt-1">
                        Tổng số hộ mới: <strong>{cand.recordB.householdCountB} hộ</strong> • Chênh lệch:{" "}
                        <span className={diff === 0 ? "text-emerald-700 font-bold" : "text-slate-800"}>
                          {diff === 0 ? "Bằng nhau (0)" : `${diff > 0 ? `+${diff}` : diff} hộ`}
                        </span>{" "}
                        • {cand.matchReason}
                      </div>
                    </div>

                    <button
                      type="button"
                      onClick={() => handleSelectCandidateManually(selectedCandidateRow.index, cand)}
                      className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs px-3 py-1.5 cursor-pointer border-0 shadow-xs whitespace-nowrap"
                    >
                      ✓ Chọn gán
                    </button>
                  </div>
                );
              })}
            </div>

            <div className="bg-slate-100 p-2.5 border-t border-slate-300 flex justify-end">
              <button
                type="button"
                onClick={() => setSelectedCandidateRow(null)}
                className="bg-slate-600 hover:bg-slate-700 text-white font-bold text-xs px-4 py-1.5 cursor-pointer border-0"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: TRA CỨU 310 XÃ SÁP NHẬP CHÍNH THỨC */}
      {showOfficialModal && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4">
          <div className="bg-white border-2 border-emerald-700 shadow-2xl max-w-3xl w-full max-h-[85vh] flex flex-col">
            <div className="bg-[#286e42] text-white p-3 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <BookOpen className="w-5 h-5 text-amber-300" />
                <div>
                  <h4 className="font-bold text-xs sm:text-sm uppercase text-white">
                    DANH MỤC 310 XÃ SÁP NHẬP ĐÃ TÍCH HỢP SẴN
                  </h4>
                  <div className="text-[11px] text-emerald-100">
                    Trích xuất đầy đủ 10 trang văn bản sáp nhập tỉnh Thái Bình, Hưng Yên...
                  </div>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowOfficialModal(false)}
                className="text-white hover:text-amber-300 font-bold text-sm bg-transparent border-0 cursor-pointer"
              >
                ✕
              </button>
            </div>

            <div className="p-3 border-b border-slate-200 bg-slate-50 flex items-center gap-2">
              <Search className="w-4 h-4 text-slate-400" />
              <input
                type="text"
                value={officialSearchQuery}
                onChange={e => setOfficialSearchQuery(e.target.value)}
                placeholder="Gõ tên xã cũ hoặc xã mới để tra cứu nhanh..."
                className="w-full text-xs border border-slate-300 px-2 py-1.5 bg-white"
                autoFocus
              />
              <span className="text-xs text-slate-500 whitespace-nowrap">
                {filteredOfficialCommunes.length} kết quả
              </span>
            </div>

            <div className="p-3 overflow-y-auto flex-1">
              <table className="w-full text-left border-collapse text-xs">
                <thead className="bg-slate-100 text-slate-700 font-bold sticky top-0 border-b border-slate-300">
                  <tr>
                    <th className="p-2 border-r border-slate-200 w-12 text-center">STT</th>
                    <th className="p-2 border-r border-slate-200">XÃ CŨ (LÚC CHƯA SÁP NHẬP)</th>
                    <th className="p-2 border-r border-slate-200 text-center w-8">➔</th>
                    <th className="p-2">XÃ MỚI (SAU SÁP NHẬP)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {filteredOfficialCommunes.map((row, rIdx) => (
                    <tr key={rIdx} className="hover:bg-emerald-50/50">
                      <td className="p-2 text-center border-r border-slate-200 text-slate-500 font-mono">
                        {rIdx + 1}
                      </td>
                      <td className="p-2 border-r border-slate-200 font-semibold text-slate-900">
                        {row.communeA}
                      </td>
                      <td className="p-2 border-r border-slate-200 text-center text-emerald-700 font-bold">
                        ➔
                      </td>
                      <td className="p-2 font-bold text-emerald-900">
                        {row.communeB}
                        {row.districtNote && (
                          <span className="ml-2 text-[10px] text-slate-500 font-normal">
                            (huyện {row.districtNote})
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="bg-slate-100 p-2.5 border-t border-slate-300 flex items-center justify-between">
              <span className="text-xs text-slate-600">
                Toàn bộ 310 xã này đang được hệ thống tự động sử dụng làm từ điển đối chiếu địa bàn.
              </span>
              <button
                type="button"
                onClick={() => setShowOfficialModal(false)}
                className="bg-emerald-800 hover:bg-emerald-900 text-white font-bold text-xs px-4 py-1.5 cursor-pointer border-0"
              >
                Đóng tra cứu
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL ĐỐI SOÁT CỤM HỘ DÂN (≥ 30 NGƯỜI TRÙNG TÊN) */}
      {showClusterResolver && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-2 sm:p-4 overflow-y-auto">
          <div className="bg-white border-2 border-slate-800 shadow-2xl max-w-6xl w-full max-h-[92vh] flex flex-col font-sans overflow-hidden">
            <HouseholdClusterResolver
              initialDataA={dataA}
              initialColumnsA={columnsA}
              initialDataB={dataB}
              initialColumnsB={columnsB}
              onApplyMergedToMain={(merged, cols) => {
                if (onApplyMergedToMain) {
                  onApplyMergedToMain(merged, cols);
                }
                setShowClusterResolver(false);
                setAppliedMsg("Đã áp dụng kết quả đối soát cụm hộ dân vào dữ liệu đang làm việc!");
              }}
              onClose={() => setShowClusterResolver(false)}
            />
          </div>
        </div>
      )}

      {/* MODAL MA TRẬN GOM 65 XÃ MỚI ➔ ĐỦ TOÀN BỘ XÃ CŨ & ĐỊA BÀN */}
      {showCommuneClusterMatrix && (
        <CommuneClusterMatrixModal
          dataA={dataA}
          colCommuneNameA={colCommuneNameA}
          colCommuneCodeA={colCommuneCodeA}
          colVillageNameA={colVillageNameA}
          dataB={dataB}
          colCommuneNameB={colCommuneNameB}
          colCommuneCodeB={colCommuneCodeB}
          colVillageNameB={colVillageNameB}
          customCommuneCrosswalk={customCommuneCrosswalk}
          onUpdateCustomCrosswalk={newList => setCustomCommuneCrosswalk(newList)}
          onSelectCommuneToFilter={communeAName => {
            setSelectedCommuneA(communeAName);
            setAppliedMsg(`📍 Đã chọn lọc bảng làm việc cho xã: "${communeAName}". Hãy bấm nút "Khởi Chạy Đối Soát Địa Bàn" để bắt đầu!`);
          }}
          onClose={() => setShowCommuneClusterMatrix(false)}
        />
      )}
    </div>
  );
};
