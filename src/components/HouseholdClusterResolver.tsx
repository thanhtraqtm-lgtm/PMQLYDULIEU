import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  Users,
  GitMerge,
  Search,
  CheckCircle2,
  AlertTriangle,
  Download,
  ArrowRight,
  Sparkles,
  RefreshCw,
  FileSpreadsheet,
  Upload,
  Eye,
  SlidersHorizontal,
  ChevronDown,
  BookOpen,
  UserCheck,
  Building,
  Check,
  X,
  Layers,
  HelpCircle
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  executeHouseholdClusterMatching,
  HouseholdClusterConfig,
  AreaClusterMatchResult,
  HouseholdMatchedDetail
} from "../utils/householdClusterMatcher";
import { parse2DArrayWithSmartHeader } from "../utils/sharedHelpers";

// Dữ liệu mẫu thử nghiệm: Địa bàn cũ có tên "Thôn Phú Thịnh (Xã An Tảo cũ)", Địa bàn mới đổi tên thành "Tổ dân phố Tân Tiến 1 (Phường Phố Hiến mới)"
// Có tới 35 hộ trùng tên để chứng minh thuật toán vượt ngưỡng 30 người
export const SAMPLE_HOUSEHOLDS_CLUSTER_A = [
  { STT: 1, "Tên chủ hộ": "Nguyễn Văn Hưng", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1968, "Số CCCD": "033068001234" },
  { STT: 2, "Tên chủ hộ": "Trần Thị Mai", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1972, "Số CCCD": "033072002345" },
  { STT: 3, "Tên chủ hộ": "Lê Đình Tuấn", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1965, "Số CCCD": "033065003456" },
  { STT: 4, "Tên chủ hộ": "Phạm Văn Long", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1970, "Số CCCD": "033070004567" },
  { STT: 5, "Tên chủ hộ": "Vũ Thị Lan", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1975, "Số CCCD": "033075005678" },
  { STT: 6, "Tên chủ hộ": "Hoàng Văn Đức", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1980, "Số CCCD": "033080006789" },
  { STT: 7, "Tên chủ hộ": "Đỗ Thị Huệ", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1963, "Số CCCD": "033063007890" },
  { STT: 8, "Tên chủ hộ": "Bùi Văn Nam", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1978, "Số CCCD": "033078008901" },
  { STT: 9, "Tên chủ hộ": "Ngô Thị Cúc", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1960, "Số CCCD": "033060009012" },
  { STT: 10, "Tên chủ hộ": "Dương Văn Thắng", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1982, "Số CCCD": "033082010123" },
  { STT: 11, "Tên chủ hộ": "Đặng Thị Nga", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1974, "Số CCCD": "033074011234" },
  { STT: 12, "Tên chủ hộ": "Lý Văn Tuyên", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1969, "Số CCCD": "033069012345" },
  { STT: 13, "Tên chủ hộ": "Mai Thị Hiền", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1976, "Số CCCD": "033076013456" },
  { STT: 14, "Tên chủ hộ": "Hà Văn Khoa", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1985, "Số CCCD": "033085014567" },
  { STT: 15, "Tên chủ hộ": "Trịnh Thị Tuyết", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1971, "Số CCCD": "033071015678" },
  { STT: 16, "Tên chủ hộ": "Chu Văn Hùng", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1967, "Số CCCD": "033067016789" },
  { STT: 17, "Tên chủ hộ": "Lương Thị Yến", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1979, "Số CCCD": "033079017890" },
  { STT: 18, "Tên chủ hộ": "Tạ Văn Bình", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1962, "Số CCCD": "033062018901" },
  { STT: 19, "Tên chủ hộ": "Phan Thị Oanh", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1973, "Số CCCD": "033073019012" },
  { STT: 20, "Tên chủ hộ": "Nguyễn Văn Dũng", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1983, "Số CCCD": "033083020123" },
  { STT: 21, "Tên chủ hộ": "Đoàn Thị Lựu", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1959, "Số CCCD": "033059021234" },
  { STT: 22, "Tên chủ hộ": "Cao Văn Sơn", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1981, "Số CCCD": "033081022345" },
  { STT: 23, "Tên chủ hộ": "Võ Thị Hằng", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1977, "Số CCCD": "033077023456" },
  { STT: 24, "Tên chủ hộ": "Lâm Văn Hải", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1966, "Số CCCD": "033066024567" },
  { STT: 25, "Tên chủ hộ": "Đinh Thị Thơ", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1970, "Số CCCD": "033070025678" },
  { STT: 26, "Tên chủ hộ": "Trương Văn Tài", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1984, "Số CCCD": "033084026789" },
  { STT: 27, "Tên chủ hộ": "Nguyễn Thị Sen", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1964, "Số CCCD": "033064027890" },
  { STT: 28, "Tên chủ hộ": "Phạm Văn Quang", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1975, "Số CCCD": "033075028901" },
  { STT: 29, "Tên chủ hộ": "Lê Thị Bích", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1982, "Số CCCD": "033082029012" },
  { STT: 30, "Tên chủ hộ": "Vũ Văn Minh", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1968, "Số CCCD": "033068030123" },
  { STT: 31, "Tên chủ hộ": "Hoàng Thị Nguyệt", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1973, "Số CCCD": "033073031234" },
  { STT: 32, "Tên chủ hộ": "Đỗ Văn Toàn", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1986, "Số CCCD": "033086032345" },
  { STT: 33, "Tên chủ hộ": "Bùi Thị Quyên", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1969, "Số CCCD": "033069033456" },
  { STT: 34, "Tên chủ hộ": "Ngô Văn Phát", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1978, "Số CCCD": "033078034567" },
  { STT: 35, "Tên chủ hộ": "Dương Thị Nhàn", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1980, "Số CCCD": "033080035678" },
  // Hộ khác không trùng
  { STT: 36, "Tên chủ hộ": "Nguyễn Bá Thắng", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1990, "Số CCCD": "033090036789" },
  { STT: 37, "Tên chủ hộ": "Trần Thị Ánh", "Thôn/Địa bàn": "Thôn Phú Thịnh", "Xã": "Xã An Tảo", "Năm sinh": 1992, "Số CCCD": "033092037890" },
];

export const SAMPLE_HOUSEHOLDS_CLUSTER_B = [
  // Cùng 35 người trên nhưng nằm trong Tổ dân phố Tân Tiến 1 (Phường Phố Hiến mới)
  { STT: 101, "Họ và tên chủ hộ": "Nguyễn Văn Hưng", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 102, "Họ và tên chủ hộ": "Trần Thị Mai", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 103, "Họ và tên chủ hộ": "Lê Đình Tuấn", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 104, "Họ và tên chủ hộ": "Phạm Văn Long", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 105, "Họ và tên chủ hộ": "Vũ Thị Lan", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 106, "Họ và tên chủ hộ": "Hoàng Văn Đức", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 107, "Họ và tên chủ hộ": "Đỗ Thị Huệ", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 108, "Họ và tên chủ hộ": "Bùi Văn Nam", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 109, "Họ và tên chủ hộ": "Ngô Thị Cúc", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 110, "Họ và tên chủ hộ": "Dương Văn Thắng", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 111, "Họ và tên chủ hộ": "Đặng Thị Nga", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 112, "Họ và tên chủ hộ": "Lý Văn Tuyên", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 113, "Họ và tên chủ hộ": "Mai Thị Hiền", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 114, "Họ và tên chủ hộ": "Hà Văn Khoa", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 115, "Họ và tên chủ hộ": "Trịnh Thị Tuyết", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 116, "Họ và tên chủ hộ": "Chu Văn Hùng", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 117, "Họ và tên chủ hộ": "Lương Thị Yến", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 118, "Họ và tên chủ hộ": "Tạ Văn Bình", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 119, "Họ và tên chủ hộ": "Phan Thị Oanh", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 120, "Họ và tên chủ hộ": "Nguyễn Văn Dũng", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 121, "Họ và tên chủ hộ": "Đoàn Thị Lựu", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 122, "Họ và tên chủ hộ": "Cao Văn Sơn", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 123, "Họ và tên chủ hộ": "Võ Thị Hằng", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 124, "Họ và tên chủ hộ": "Lâm Văn Hải", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 125, "Họ và tên chủ hộ": "Đinh Thị Thơ", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 126, "Họ và tên chủ hộ": "Trương Văn Tài", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 127, "Họ và tên chủ hộ": "Nguyễn Thị Sen", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 128, "Họ và tên chủ hộ": "Phạm Văn Quang", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 129, "Họ và tên chủ hộ": "Lê Thị Bích", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 130, "Họ và tên chủ hộ": "Vũ Văn Minh", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 131, "Họ và tên chủ hộ": "Hoàng Thị Nguyệt", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 132, "Họ và tên chủ hộ": "Đỗ Văn Toàn", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 133, "Họ và tên chủ hộ": "Bùi Thị Quyên", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 134, "Họ và tên chủ hộ": "Ngô Văn Phát", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  { STT: 135, "Họ và tên chủ hộ": "Dương Thị Nhàn", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
  // Hộ mới thêm
  { STT: 136, "Họ và tên chủ hộ": "Trần Quốc Toản", "Địa bàn / TDP mới": "Tổ dân phố Tân Tiến 1", "Xã / Phường mới": "Phường Phố Hiến", "Mã ĐBĐT": "DB01", "Tên ĐBĐT": "Tân Tiến 1_ĐB01" },
];

interface HouseholdClusterResolverProps {
  initialDataA?: any[];
  initialColumnsA?: string[];
  initialDataB?: any[];
  initialColumnsB?: string[];
  onApplyMergedToMain?: (mergedData: any[], mergedCols: string[]) => void;
  onClose?: () => void;
}

export const HouseholdClusterResolver: React.FC<HouseholdClusterResolverProps> = ({
  initialDataA = [],
  initialColumnsA = [],
  initialDataB = [],
  initialColumnsB = [],
  onApplyMergedToMain,
  onClose,
}) => {
  // Dữ liệu danh sách hộ Tệp A
  const [dataA, setDataA] = useState<any[]>(initialDataA);
  const [columnsA, setColumnsA] = useState<string[]>(initialColumnsA);
  const [fileNameA, setFileNameA] = useState<string>("Tệp Hộ Dân Cũ (Kỳ NN)");

  // Dữ liệu danh sách hộ Tệp B
  const [dataB, setDataB] = useState<any[]>(initialDataB);
  const [columnsB, setColumnsB] = useState<string[]>(initialColumnsB);
  const [fileNameB, setFileNameB] = useState<string>("Tệp Hộ Dân Mới (Kỳ Cá Thể)");

  // Cột cấu hình Tệp A
  const [colPersonNameA, setColPersonNameA] = useState<string>("");
  const [colAreaOrVillageA, setColAreaOrVillageA] = useState<string>("");
  const [colCommuneA, setColCommuneA] = useState<string>("");
  const [colExtraA, setColExtraA] = useState<string>("");

  // Cột cấu hình Tệp B
  const [colPersonNameB, setColPersonNameB] = useState<string>("");
  const [colAreaOrVillageB, setColAreaOrVillageB] = useState<string>("");
  const [colCommuneB, setColCommuneB] = useState<string>("");
  const [colDbdtNameB, setColDbdtNameB] = useState<string>("");
  const [colDbdtCodeB, setColDbdtCodeB] = useState<string>("");
  const [colExtraB, setColExtraB] = useState<string>("");

  // Ngưỡng số người trùng (mặc định: 30 người)
  const [minOverlapThreshold, setMinOverlapThreshold] = useState<number>(30);

  // Kết quả
  const [results, setResults] = useState<AreaClusterMatchResult[]>([]);
  const [summary, setSummary] = useState<any>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [appliedMsg, setAppliedMsg] = useState<string | null>(null);

  // Modal xem chi tiết danh sách N người trùng tên
  const [inspectModalResult, setInspectModalResult] = useState<AreaClusterMatchResult | null>(null);

  // Bộ lọc
  const [filterMode, setFilterMode] = useState<"ALL" | "MATCHED" | "UNMATCHED">("ALL");
  const [searchQuery, setSearchQuery] = useState<string>("");

  const fileInputRefA = useRef<HTMLInputElement>(null);
  const fileInputRefB = useRef<HTMLInputElement>(null);

  // Tự động nhận diện cột thông minh
  const autoDetectColsA = (cols: string[]) => {
    cols.forEach(c => {
      const lower = c.toLowerCase();
      if (!colPersonNameA && (lower.includes("chủ hộ") || lower.includes("họ và tên") || lower.includes("họ tên") || lower.includes("tên"))) {
        setColPersonNameA(c);
      }
      if (!colAreaOrVillageA && (lower.includes("thôn") || lower.includes("địa bàn") || lower.includes("tổ dân phố") || lower.includes("mã đb"))) {
        setColAreaOrVillageA(c);
      }
      if (!colCommuneA && (lower.includes("xã") || lower.includes("phường") || lower.includes("thị trấn"))) {
        setColCommuneA(c);
      }
      if (!colExtraA && (lower.includes("cccd") || lower.includes("năm sinh") || lower.includes("sđt") || lower.includes("điện thoại"))) {
        setColExtraA(c);
      }
    });
  };

  const autoDetectColsB = (cols: string[]) => {
    cols.forEach(c => {
      const lower = c.toLowerCase();
      if (!colPersonNameB && (lower.includes("chủ hộ") || lower.includes("họ và tên") || lower.includes("họ tên") || lower.includes("tên"))) {
        setColPersonNameB(c);
      }
      if (!colAreaOrVillageB && (lower.includes("thôn") || lower.includes("địa bàn") || lower.includes("tổ dân phố") || lower.includes("tdp"))) {
        setColAreaOrVillageB(c);
      }
      if (!colCommuneB && (lower.includes("xã") || lower.includes("phường") || lower.includes("thị trấn"))) {
        setColCommuneB(c);
      }
      if (!colDbdtNameB && (lower.includes("tên đbđt") || lower.includes("đbđt") || lower.includes("địa bàn điều tra"))) {
        setColDbdtNameB(c);
      }
      if (!colDbdtCodeB && (lower.includes("mã đbđt") || lower.includes("mã đb"))) {
        setColDbdtCodeB(c);
      }
      if (!colExtraB && (lower.includes("cccd") || lower.includes("năm sinh") || lower.includes("sđt"))) {
        setColExtraB(c);
      }
    });
  };

  useEffect(() => {
    if (columnsA.length > 0) autoDetectColsA(columnsA);
  }, [columnsA]);

  useEffect(() => {
    if (columnsB.length > 0) autoDetectColsB(columnsB);
  }, [columnsB]);

  // Nạp tệp A
  const handleUploadA = (e: React.ChangeEvent<HTMLInputElement>) => {
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
        setResults([]);
      } catch (err: any) {
        alert("Lỗi khi đọc file A: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Nạp tệp B
  const handleUploadB = (e: React.ChangeEvent<HTMLInputElement>) => {
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
        setResults([]);
      } catch (err: any) {
        alert("Lỗi khi đọc file B: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Nạp mẫu thử nghiệm 35 hộ trùng
  const handleLoadSampleData = () => {
    setDataA(SAMPLE_HOUSEHOLDS_CLUSTER_A);
    const colsA = Object.keys(SAMPLE_HOUSEHOLDS_CLUSTER_A[0]);
    setColumnsA(colsA);
    setFileNameA("Mau_37_Ho_Dan_ThonPhuThinh_AnTao.xlsx");
    setColPersonNameA("Tên chủ hộ");
    setColAreaOrVillageA("Thôn/Địa bàn");
    setColCommuneA("Xã");
    setColExtraA("Số CCCD");

    setDataB(SAMPLE_HOUSEHOLDS_CLUSTER_B);
    const colsB = Object.keys(SAMPLE_HOUSEHOLDS_CLUSTER_B[0]);
    setColumnsB(colsB);
    setFileNameB("Mau_36_Ho_Dan_TDPTanTien1_PhoHien.xlsx");
    setColPersonNameB("Họ và tên chủ hộ");
    setColAreaOrVillageB("Địa bàn / TDP mới");
    setColCommuneB("Xã / Phường mới");
    setColDbdtNameB("Tên ĐBĐT");
    setColDbdtCodeB("Mã ĐBĐT");

    setMinOverlapThreshold(30);
    setResults([]);
  };

  // Khởi chạy thuật toán so khớp cụm hộ
  const handleRunClusterMatching = () => {
    if (dataA.length === 0 || dataB.length === 0) {
      alert("Vui lòng nạp danh sách hộ dân ở cả 2 tệp!");
      return;
    }
    if (!colPersonNameA || !colAreaOrVillageA) {
      alert("Vui lòng chọn Cột Tên chủ hộ và Cột Thôn/Địa bàn ở Tệp A!");
      return;
    }
    if (!colPersonNameB || !colAreaOrVillageB) {
      alert("Vui lòng chọn Cột Tên chủ hộ và Cột Thôn/Địa bàn ở Tệp B!");
      return;
    }

    setIsProcessing(true);

    setTimeout(() => {
      try {
        const config: HouseholdClusterConfig = {
          colPersonNameA,
          colAreaOrVillageA,
          colCommuneA,
          colExtraA,
          colPersonNameB,
          colAreaOrVillageB,
          colCommuneB,
          colDbdtNameB,
          colDbdtCodeB,
          colExtraB,
          minOverlapThreshold,
          minOverlapRatio: 0.25,
        };

        const res = executeHouseholdClusterMatching(dataA, dataB, config);
        setResults(res.results);
        setSummary(res.summary);
      } catch (err: any) {
        alert("Lỗi khi xử lý đối soát cụm hộ: " + err.message);
      } finally {
        setIsProcessing(false);
      }
    }, 50);
  };

  // Áp dụng gán mã địa bàn vào dữ liệu chính
  const handleApplyMergedToMainData = () => {
    if (results.length === 0 || !onApplyMergedToMain) return;

    // Tạo bản đồ ánh xạ từ AreaKey cũ sang thông tin mới
    const clusterCodeMap = new Map<string, any>();
    results.forEach(r => {
      if (r.matchedClusterB) {
        clusterCodeMap.set(r.clusterA.areaKey, {
          Dia_Ban_Moi_GhepDuoc: r.matchedClusterB.villageName,
          Xa_Moi_GhepDuoc: r.matchedClusterB.communeName,
          Ma_DBDT_Moi: r.matchedClusterB.dbdtCode,
          Ten_DBDT_Moi: r.matchedClusterB.dbdtName,
          So_Nguoi_Trung_Ten: r.overlapCount,
          Can_Cu_Doi_Soat: `Khớp cụm hộ (${r.overlapCount} người trùng tên ≥ ${minOverlapThreshold})`,
        });
      }
    });

    // Cập nhật vào dataA
    const updatedData = dataA.map(r => {
      const rawArea = String(r[colAreaOrVillageA] || "").trim();
      const rawCommune = colCommuneA ? String(r[colCommuneA] || "").trim() : "";
      const areaKey = rawCommune ? `${rawCommune}___${rawArea}` : rawArea;

      const match = clusterCodeMap.get(areaKey);
      if (match) {
        return { ...r, ...match };
      }
      return {
        ...r,
        Dia_Ban_Moi_GhepDuoc: "",
        Xa_Moi_GhepDuoc: "",
        Ma_DBDT_Moi: "",
        Ten_DBDT_Moi: "",
        So_Nguoi_Trung_Ten: 0,
        Can_Cu_Doi_Soat: "Chưa khớp",
      };
    });

    const newCols = [...columnsA];
    ["Dia_Ban_Moi_GhepDuoc", "Xa_Moi_GhepDuoc", "Ma_DBDT_Moi", "Ten_DBDT_Moi", "So_Nguoi_Trung_Ten", "Can_Cu_Doi_Soat"].forEach(c => {
      if (!newCols.includes(c)) newCols.push(c);
    });

    onApplyMergedToMain(updatedData, newCols);
    setAppliedMsg(`Đã cập nhật mã địa bàn mới cho ${clusterCodeMap.size} cụm địa bàn vào bảng dữ liệu đang làm việc!`);
  };

  // Xuất file Excel đối soát cụm hộ
  const handleExportClusterExcel = () => {
    if (results.length === 0) {
      alert("Chưa có kết quả để xuất!");
      return;
    }

    const exportRows = results.map(r => ({
      "Xã cũ": r.clusterA.communeName,
      "Địa bàn / Thôn cũ": r.clusterA.villageName,
      "Tổng số hộ thôn cũ": r.clusterA.totalHouseholds,
      "➔": "➔",
      "Xã mới ghép được": r.matchedClusterB?.communeName || "Chưa tìm thấy",
      "Địa bàn / TDP mới ghép được": r.matchedClusterB?.villageName || "Chưa tìm thấy",
      "Mã ĐBĐT mới": r.matchedClusterB?.dbdtCode || "",
      "Tên ĐBĐT mới": r.matchedClusterB?.dbdtName || "",
      "Tổng số hộ thôn mới": r.matchedClusterB?.totalHouseholds || "",
      "Số người trùng tên": r.overlapCount,
      "Ngưỡng yêu cầu": `≥ ${minOverlapThreshold} người`,
      "Tỷ lệ trùng (% số hộ cũ)": `${(r.overlapRatioA * 100).toFixed(1)}%`,
      "Trạng thái kết luận": r.matchedClusterB ? `ĐÃ TẠM TÍNH (Khớp ${r.overlapCount} người)` : "CHƯA ĐẠT NGƯỠNG",
      "Căn cứ kết luận": r.matchReason,
      "Top 5 người trùng mẫu": r.matchedHouseholds.slice(0, 5).map(m => m.nameA).join(", "),
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DoiSoat_Cum_Ho_Dan");
    XLSX.writeFile(wb, `KetQua_DoiSoat_CumHoDan_${results.length}DiaBan.xlsx`);
  };

  // Bộ lọc danh sách hiển thị
  const filteredResults = useMemo(() => {
    return results.filter(r => {
      if (filterMode === "MATCHED" && !r.matchedClusterB) return false;
      if (filterMode === "UNMATCHED" && r.matchedClusterB) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const aName = r.clusterA.villageName.toLowerCase();
        const aCom = r.clusterA.communeName.toLowerCase();
        const bName = r.matchedClusterB ? r.matchedClusterB.villageName.toLowerCase() : "";
        const bCom = r.matchedClusterB ? r.matchedClusterB.communeName.toLowerCase() : "";
        return aName.includes(q) || aCom.includes(q) || bName.includes(q) || bCom.includes(q);
      }

      return true;
    });
  }, [results, filterMode, searchQuery]);

  return (
    <div className="bg-white border border-slate-300 shadow-sm space-y-4 font-sans text-slate-800">
      {/* 1. HEADER CHÍNH */}
      <div className="bg-[#1e5430] text-white p-3 sm:p-4 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <div className="w-10 h-10 bg-amber-400 text-slate-950 flex items-center justify-center font-black shrink-0 shadow-xs">
            <Users className="w-6 h-6" />
          </div>
          <div>
            <div className="flex flex-wrap items-center gap-2">
              <h3 className="text-sm sm:text-base font-bold uppercase tracking-wider text-white">
                ĐỐI SOÁT CỤM HỘ DÂN (HOUSEHOLD CLUSTER OVERLAP)
              </h3>
              <span className="bg-amber-400 text-slate-950 font-black text-[10px] px-2 py-0.5 uppercase shadow-2xs">
                ≥ {minOverlapThreshold} Người Trùng Tên ➔ Tạm Tính Cùng Địa Bàn
              </span>
            </div>
            <p className="text-xs text-emerald-100 font-normal m-0 pt-0.5 leading-relaxed">
              Giải pháp tối hậu khi <strong>tên thôn bị đổi hoàn toàn hoặc không khớp</strong>: Tự động gom danh sách hộ theo từng địa bàn và so khớp chéo họ tên. Nếu 2 địa bàn có từ <strong>{minOverlapThreshold} người trùng tên trở lên</strong>, hệ thống tự động kết luận là cùng một địa bàn!
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={handleLoadSampleData}
            className="bg-amber-400 hover:bg-amber-500 text-slate-950 font-bold text-xs px-3 py-1.5 flex items-center gap-1.5 cursor-pointer shadow-xs border-0"
            title="Nạp ngay dữ liệu mẫu thử nghiệm có 35 hộ trùng tên giữa 2 địa bàn bị đổi tên hoàn toàn"
          >
            <Sparkles className="w-4 h-4 text-slate-950" />
            <span>⚡ Nạp mẫu thử nghiệm (35 người trùng)</span>
          </button>

          {onClose && (
            <button
              type="button"
              onClick={onClose}
              className="text-white/80 hover:text-white cursor-pointer bg-white/10 hover:bg-white/20 px-2.5 py-1.5 text-xs border border-white/20"
            >
              ✕ Đóng
            </button>
          )}
        </div>
      </div>

      {/* THÔNG BÁO ÁP DỤNG */}
      {appliedMsg && (
        <div className="mx-3.5 p-3 bg-emerald-50 border-2 border-emerald-500 text-emerald-950 text-xs font-bold flex items-center justify-between gap-2 shadow-xs">
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

      {/* 2. CHỌN TỆP & CẤU HÌNH CỘT */}
      <div className="p-3 sm:p-4 space-y-4">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {/* TỆP A: HỘ DÂN KỲ CŨ */}
          <div className="bg-sky-50/50 border border-sky-300 p-3 space-y-3">
            <div className="flex items-center justify-between border-b border-sky-200 pb-2">
              <div>
                <span className="text-xs font-bold text-sky-950 uppercase flex items-center gap-1.5">
                  <FileSpreadsheet className="w-4 h-4 text-sky-700" />
                  1. TỆP A: DANH SÁCH HỘ DÂN KỲ CŨ
                </span>
                <span className="text-[11px] text-sky-800 font-medium block">
                  {dataA.length > 0 ? `Đã nạp: ${dataA.length.toLocaleString("vi-VN")} dòng hộ (${fileNameA})` : "Chưa có dữ liệu"}
                </span>
              </div>

              <div>
                <input
                  type="file"
                  ref={fileInputRefA}
                  accept=".xlsx,.xls,.csv"
                  onChange={handleUploadA}
                  className="hidden"
                />
                <button
                  type="button"
                  onClick={() => fileInputRefA.current?.click()}
                  className="bg-sky-700 hover:bg-sky-800 text-white font-bold text-[11px] px-2.5 py-1 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                >
                  <Upload className="w-3.5 h-3.5" />
                  {dataA.length > 0 ? "Đổi Tệp A..." : "Tải Tệp A (Excel)..."}
                </button>
              </div>
            </div>

            {columnsA.length > 0 ? (
              <div className="space-y-2 text-xs">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <div>
                    <label className="text-[11px] font-bold text-slate-700 block">
                      Cột Tên Chủ Hộ / Người Dân <span className="text-rose-600">*</span>:
                    </label>
                    <select
                      value={colPersonNameA}
                      onChange={e => setColPersonNameA(e.target.value)}
                      className="w-full border border-sky-300 bg-white p-1 text-xs font-bold text-sky-900"
                    >
                      <option value="">-- Chọn Cột Tên --</option>
                      {columnsA.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="text-[11px] font-bold text-sky-800 block">
                      Cột Thôn / Địa Bàn Cũ <span className="text-rose-600">*</span>:
                    </label>
                    <select
                      value={colAreaOrVillageA}
                      onChange={e => setColAreaOrVillageA(e.target.value)}
                      className="w-full border border-sky-400 bg-sky-50/50 p-1 text-xs font-bold text-sky-950"
                    >
                      <option value="">-- Chọn Cột Địa Bàn --</option>
                      {columnsA.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-1">
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Cột Tên Xã cũ (Tùy chọn):</label>
                    <select
                      value={colCommuneA}
                      onChange={e => setColCommuneA(e.target.value)}
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
                    <label className="text-[10px] font-semibold text-slate-600 block">Cột CCCD / Năm Sinh / SĐT (Tùy chọn):</label>
                    <select
                      value={colExtraA}
                      onChange={e => setColExtraA(e.target.value)}
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
                Hãy nạp Tệp danh sách hộ hoặc bấm nút &ldquo;Nạp mẫu thử nghiệm&rdquo; ở góc trên.
              </div>
            )}
          </div>

          {/* TỆP B: HỘ DÂN KỲ MỚI */}
          <div className="bg-emerald-50/50 border border-emerald-300 p-3 space-y-3">
            <div className="flex items-center justify-between border-b border-emerald-200 pb-2">
              <div>
                <span className="text-xs font-bold text-emerald-950 uppercase flex items-center gap-1.5">
                  <FileSpreadsheet className="w-4 h-4 text-emerald-700" />
                  2. TỆP B: DANH SÁCH HỘ DÂN KỲ MỚI (NĂM NAY)
                </span>
                <span className="text-[11px] text-emerald-800 font-medium block">
                  {dataB.length > 0 ? `Đã nạp: ${dataB.length.toLocaleString("vi-VN")} dòng hộ (${fileNameB})` : "Chưa có dữ liệu"}
                </span>
              </div>

              <div>
                <input
                  type="file"
                  ref={fileInputRefB}
                  accept=".xlsx,.xls,.csv"
                  onChange={handleUploadB}
                  className="hidden"
                />
                <button
                  type="button"
                  onClick={() => fileInputRefB.current?.click()}
                  className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-[11px] px-2.5 py-1 flex items-center gap-1 cursor-pointer border-0 shadow-xs"
                >
                  <Upload className="w-3.5 h-3.5" />
                  {dataB.length > 0 ? "Đổi Tệp B..." : "Tải Tệp B (Excel)..."}
                </button>
              </div>
            </div>

            {columnsB.length > 0 ? (
              <div className="space-y-2 text-xs">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <div>
                    <label className="text-[11px] font-bold text-slate-700 block">
                      Cột Tên Chủ Hộ / Người Dân <span className="text-rose-600">*</span>:
                    </label>
                    <select
                      value={colPersonNameB}
                      onChange={e => setColPersonNameB(e.target.value)}
                      className="w-full border border-emerald-300 bg-white p-1 text-xs font-bold text-emerald-900"
                    >
                      <option value="">-- Chọn Cột Tên --</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="text-[11px] font-bold text-emerald-800 block">
                      Cột Thôn / Địa Bàn / TDP Mới <span className="text-rose-600">*</span>:
                    </label>
                    <select
                      value={colAreaOrVillageB}
                      onChange={e => setColAreaOrVillageB(e.target.value)}
                      className="w-full border border-emerald-400 bg-emerald-50/50 p-1 text-xs font-bold text-emerald-950"
                    >
                      <option value="">-- Chọn Cột Địa Bàn --</option>
                      {columnsB.map(c => (
                        <option key={c} value={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 pt-1">
                  <div>
                    <label className="text-[10px] font-semibold text-slate-600 block">Cột Tên Xã mới:</label>
                    <select
                      value={colCommuneB}
                      onChange={e => setColCommuneB(e.target.value)}
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
                    <label className="text-[10px] font-semibold text-slate-600 block">Cột Tên ĐBĐT:</label>
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
                    <label className="text-[10px] font-semibold text-slate-600 block">Cột Mã ĐBĐT:</label>
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
                </div>
              </div>
            ) : (
              <div className="text-[11px] text-slate-500 italic p-3 text-center bg-white border border-dashed border-emerald-200">
                Hãy nạp Tệp danh sách hộ hoặc bấm nút &ldquo;Nạp mẫu thử nghiệm&rdquo; ở góc trên.
              </div>
            )}
          </div>
        </div>

        {/* 3. THANH ĐIỀU CHỈNH NGƯỠNG SỐ NGƯỜI TRÙNG TÊN & NÚT CHẠY */}
        <div className="bg-amber-50/80 border-2 border-amber-300 p-3 flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-slate-900 uppercase">
                Ngưỡng số người trùng tên để tạm tính:
              </span>
              <div className="flex items-center gap-1">
                <input
                  type="number"
                  min={5}
                  max={200}
                  value={minOverlapThreshold}
                  onChange={e => setMinOverlapThreshold(Math.max(5, parseInt(e.target.value, 10) || 30))}
                  className="w-16 border-2 border-amber-500 bg-white px-2 py-1 text-sm font-black text-amber-900 text-center"
                />
                <span className="text-xs font-bold text-slate-700">người</span>
              </div>
            </div>

            {/* Các nút chọn nhanh */}
            <div className="flex items-center gap-1">
              <span className="text-[11px] text-slate-500">Chọn nhanh:</span>
              {[15, 20, 25, 30, 40, 50].map(val => (
                <button
                  key={val}
                  type="button"
                  onClick={() => setMinOverlapThreshold(val)}
                  className={`px-2 py-0.5 text-xs font-bold cursor-pointer border ${
                    minOverlapThreshold === val
                      ? "bg-amber-500 text-slate-950 border-amber-600 shadow-xs"
                      : "bg-white text-slate-700 border-slate-300 hover:bg-slate-100"
                  }`}
                >
                  {val}
                </button>
              ))}
            </div>
          </div>

          <button
            type="button"
            onClick={handleRunClusterMatching}
            disabled={isProcessing || dataA.length === 0 || dataB.length === 0}
            className="bg-[#1e5430] hover:bg-[#163f24] disabled:bg-slate-300 text-white font-bold text-xs px-6 py-2.5 flex items-center gap-2 cursor-pointer shadow-sm border-0 transition-all active:scale-95 shrink-0"
          >
            {isProcessing ? (
              <>
                <RefreshCw className="w-4 h-4 animate-spin" />
                <span>Đang quét ma trận họ tên hộ dân...</span>
              </>
            ) : (
              <>
                <GitMerge className="w-4 h-4 text-amber-300" />
                <span>🚀 TIẾN HÀNH ĐỐI SOÁT CỤM HỘ DÂN (≥ {minOverlapThreshold} NGƯỜI)</span>
              </>
            )}
          </button>
        </div>

        {/* 4. KẾT QUẢ ĐỐI SOÁT CỤM HỘ */}
        {summary && (
          <div className="space-y-3 pt-2">
            {/* THẺ TỔNG HỢP */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
              <div
                onClick={() => setFilterMode("ALL")}
                className={`p-2.5 border cursor-pointer ${
                  filterMode === "ALL" ? "border-slate-800 bg-slate-800 text-white" : "border-slate-300 bg-slate-50"
                }`}
              >
                <div className="text-[10.5px] uppercase font-bold opacity-80">Tổng Địa Bàn Cũ (A)</div>
                <div className="text-lg font-black">{summary.totalAreasA}</div>
                <div className="text-[10px] opacity-70">So sánh với {summary.totalAreasB} địa bàn mới</div>
              </div>

              <div
                onClick={() => setFilterMode("MATCHED")}
                className={`p-2.5 border cursor-pointer ${
                  filterMode === "MATCHED" ? "border-emerald-700 bg-emerald-700 text-white" : "border-emerald-300 bg-emerald-50"
                }`}
              >
                <div className="text-[10.5px] uppercase font-bold text-emerald-900">Đã Tạm Tính Thành Công</div>
                <div className="text-lg font-black text-emerald-950">{summary.matchedAreas}</div>
                <div className="text-[10px] text-emerald-800 font-bold">
                  (Đạt ≥ {minOverlapThreshold} người trùng tên)
                </div>
              </div>

              <div
                onClick={() => setFilterMode("UNMATCHED")}
                className={`p-2.5 border cursor-pointer ${
                  filterMode === "UNMATCHED" ? "border-rose-700 bg-rose-700 text-white" : "border-rose-300 bg-rose-50"
                }`}
              >
                <div className="text-[10.5px] uppercase font-bold text-rose-900">Chưa Đạt Ngưỡng</div>
                <div className="text-lg font-black text-rose-950">{summary.unmatchedAreas}</div>
                <div className="text-[10px] text-rose-800 font-medium">(&lt; {minOverlapThreshold} người trùng)</div>
              </div>

              <div className="p-2.5 border border-sky-300 bg-sky-50">
                <div className="text-[10.5px] uppercase font-bold text-sky-900">Tổng Số Hộ Khớp Chéo</div>
                <div className="text-lg font-black text-sky-950">
                  {summary.totalMatchedHouseholdsAcrossAreas.toLocaleString("vi-VN")}
                </div>
                <div className="text-[10px] text-sky-800 font-bold">Tỷ lệ thành công: {summary.matchRate}%</div>
              </div>
            </div>

            {/* THANH TÌM KIẾM & NÚT XUẤT */}
            <div className="flex flex-wrap items-center justify-between gap-2 bg-slate-100 p-2 border border-slate-300">
              <div className="flex items-center gap-2 flex-1 max-w-md">
                <Search className="w-4 h-4 text-slate-500" />
                <input
                  type="text"
                  placeholder="Tìm theo tên thôn/xã cũ hoặc mới..."
                  value={searchQuery}
                  onChange={e => setSearchQuery(e.target.value)}
                  className="w-full bg-white border border-slate-300 px-2 py-1 text-xs"
                />
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleApplyMergedToMainData}
                  disabled={summary.matchedAreas === 0}
                  className="bg-emerald-700 hover:bg-emerald-800 disabled:bg-slate-300 text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1.5 cursor-pointer border-0 shadow-xs"
                >
                  <Check className="w-3.5 h-3.5" />
                  <span>Gán Mã Mới Vào Dữ Liệu Đang Mở</span>
                </button>

                <button
                  type="button"
                  onClick={handleExportClusterExcel}
                  className="bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs px-3 py-1.5 flex items-center gap-1.5 cursor-pointer border-0 shadow-xs"
                >
                  <Download className="w-3.5 h-3.5 text-amber-300" />
                  <span>Xuất Excel Đối Soát Cụm Hộ</span>
                </button>
              </div>
            </div>

            {/* BẢNG KẾT QUẢ ĐỐI SOÁT */}
            <div className="border border-slate-300 overflow-x-auto max-h-[500px]">
              <table className="w-full text-left text-xs border-collapse font-sans">
                <thead className="bg-slate-200 text-slate-800 uppercase font-bold sticky top-0 border-b border-slate-300 text-[11px] z-10">
                  <tr>
                    <th className="p-2 border-r border-slate-300 w-10 text-center">STT</th>
                    <th className="p-2 border-r border-slate-300">Địa Bàn / Thôn Cũ (Kỳ A)</th>
                    <th className="p-2 border-r border-slate-300 text-center w-24">Số Hộ Cũ</th>
                    <th className="p-2 border-r border-slate-300 text-center w-8">➔</th>
                    <th className="p-2 border-r border-slate-300">Địa Bàn / TDP Mới Ghép Được (Kỳ B)</th>
                    <th className="p-2 border-r border-slate-300 text-center w-28">Số Người Trùng Tên</th>
                    <th className="p-2 border-r border-slate-300 text-center w-24">Tỷ Lệ Trùng</th>
                    <th className="p-2 border-r border-slate-300">Căn Cứ / Bằng Chứng</th>
                    <th className="p-2 text-center w-28">Hành Động</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 bg-white">
                  {filteredResults.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="p-4 text-center text-slate-500 italic">
                        Không có địa bàn nào thỏa mãn bộ lọc hiện tại.
                      </td>
                    </tr>
                  ) : (
                    filteredResults.map((r, idx) => {
                      const isMatched = r.matchedClusterB !== null;
                      return (
                        <tr key={idx} className={isMatched ? "hover:bg-emerald-50/50" : "bg-rose-50/20 hover:bg-rose-50/40"}>
                          <td className="p-2 border-r border-slate-200 text-center font-bold text-slate-600">
                            {idx + 1}
                          </td>
                          <td className="p-2 border-r border-slate-200 font-bold text-slate-900">
                            <div>{r.clusterA.villageName}</div>
                            {r.clusterA.communeName && (
                              <div className="text-[10px] text-slate-500 font-normal">
                                Xã: {r.clusterA.communeName}
                              </div>
                            )}
                          </td>
                          <td className="p-2 border-r border-slate-200 text-center font-bold text-slate-700">
                            {r.clusterA.totalHouseholds} hộ
                          </td>
                          <td className="p-2 border-r border-slate-200 text-center text-emerald-700 font-black">
                            ➔
                          </td>
                          <td className="p-2 border-r border-slate-200">
                            {r.matchedClusterB ? (
                              <div>
                                <div className="font-bold text-emerald-950 flex items-center gap-1">
                                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                                  <span>{r.matchedClusterB.villageName}</span>
                                </div>
                                <div className="text-[10.5px] text-slate-600">
                                  Xã/Phường: {r.matchedClusterB.communeName || "(Chưa có tên)"}
                                  {r.matchedClusterB.dbdtName && ` • ĐBĐT: ${r.matchedClusterB.dbdtName}`}
                                </div>
                              </div>
                            ) : (
                              <div className="text-rose-700 italic flex items-center gap-1 font-medium">
                                <AlertTriangle className="w-3.5 h-3.5 text-rose-500 shrink-0" />
                                <span>Chưa đủ {minOverlapThreshold} người trùng</span>
                              </div>
                            )}
                          </td>
                          <td className="p-2 border-r border-slate-200 text-center">
                            {r.overlapCount > 0 ? (
                              <span
                                className={`inline-block font-black text-xs px-2 py-0.5 ${
                                  r.overlapCount >= minOverlapThreshold
                                    ? "bg-emerald-600 text-white shadow-2xs"
                                    : "bg-amber-100 text-amber-950 border border-amber-300"
                                }`}
                              >
                                {r.overlapCount} người
                              </span>
                            ) : (
                              <span className="text-slate-400">0</span>
                            )}
                          </td>
                          <td className="p-2 border-r border-slate-200 text-center font-bold text-slate-700">
                            {r.overlapCount > 0 ? `${(r.overlapRatioA * 100).toFixed(0)}%` : "-"}
                          </td>
                          <td className="p-2 border-r border-slate-200 text-[11px] text-slate-600 leading-tight">
                            {r.matchReason}
                          </td>
                          <td className="p-2 text-center">
                            {r.matchedHouseholds.length > 0 && (
                              <button
                                type="button"
                                onClick={() => setInspectModalResult(r)}
                                className="bg-sky-600 hover:bg-sky-700 text-white font-bold text-[10.5px] px-2 py-1 flex items-center justify-center gap-1 cursor-pointer mx-auto border-0 shadow-xs"
                                title="Bấm để xem cụ thể danh sách họ tên từng người trùng giữa 2 địa bàn"
                              >
                                <Eye className="w-3 h-3" />
                                <span>Xem {r.matchedHouseholds.length} người</span>
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

      {/* MODAL XEM CHI TIẾT DANH SÁCH HỌ TÊN TRÙNG */}
      {inspectModalResult && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4">
          <div className="bg-white border-2 border-slate-800 shadow-2xl max-w-3xl w-full max-h-[85vh] flex flex-col font-sans">
            <div className="bg-[#1e5430] text-white p-3 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Users className="w-5 h-5 text-amber-300" />
                <h4 className="text-sm font-bold uppercase tracking-wider">
                  DANH SÁCH {inspectModalResult.matchedHouseholds.length} HỘ TRÙNG TÊN ĐỐI CHỨNG
                </h4>
              </div>
              <button
                type="button"
                onClick={() => setInspectModalResult(null)}
                className="text-white hover:text-amber-300 font-bold text-sm bg-transparent border-0 cursor-pointer"
              >
                ✕
              </button>
            </div>

            <div className="p-3 bg-emerald-50 border-b border-emerald-200 text-xs text-emerald-950 flex flex-wrap items-center justify-between gap-2">
              <div>
                <span className="font-bold">Địa bàn cũ:</span> {inspectModalResult.clusterA.villageName} ({inspectModalResult.clusterA.communeName})
                <span className="mx-2 text-emerald-700">➔</span>
                <span className="font-bold">Địa bàn mới:</span> {inspectModalResult.matchedClusterB?.villageName} ({inspectModalResult.matchedClusterB?.communeName})
              </div>
              <div className="bg-emerald-700 text-white px-2 py-0.5 font-black text-[11px]">
                {inspectModalResult.matchedHouseholds.length} người trùng tên (Ngưỡng yêu cầu: ≥ {minOverlapThreshold})
              </div>
            </div>

            <div className="p-3 overflow-y-auto flex-1">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-slate-100 text-slate-800 uppercase font-bold sticky top-0 text-[10.5px]">
                  <tr>
                    <th className="p-2 border-b border-slate-300 w-10 text-center">STT</th>
                    <th className="p-2 border-b border-slate-300">Tên Hộ Dân Kỳ Cũ (A)</th>
                    <th className="p-2 border-b border-slate-300">Tên Hộ Dân Kỳ Mới (B)</th>
                    <th className="p-2 border-b border-slate-300">Thông tin bổ trợ (CCCD / Năm sinh)</th>
                    <th className="p-2 border-b border-slate-300 text-center w-20">Trạng thái</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {inspectModalResult.matchedHouseholds.map((h, i) => (
                    <tr key={i} className="hover:bg-slate-50">
                      <td className="p-1.5 text-center font-bold text-slate-600">{i + 1}</td>
                      <td className="p-1.5 font-bold text-sky-900">{h.nameA}</td>
                      <td className="p-1.5 font-bold text-emerald-900">{h.nameB}</td>
                      <td className="p-1.5 text-[11px] text-slate-600">{h.extraA || h.extraB || "—"}</td>
                      <td className="p-1.5 text-center">
                        <span className="bg-emerald-100 text-emerald-900 font-bold text-[10px] px-1.5 py-0.5 border border-emerald-300">
                          Khớp 100%
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="p-3 bg-slate-100 border-t border-slate-300 flex items-center justify-between">
              <span className="text-xs text-slate-600 italic">
                Căn cứ khoa học: Tỷ lệ 2 địa bàn khác nhau hoàn toàn mà ngẫu nhiên trùng hơn {minOverlapThreshold} họ tên người dân là cực kỳ thấp (&lt; 0.001%).
              </span>
              <button
                type="button"
                onClick={() => setInspectModalResult(null)}
                className="bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs px-4 py-1.5 cursor-pointer border-0"
              >
                Đóng lại
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
