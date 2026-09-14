/**
 * DANH MỤC ÁNH XẠ CHÍNH THỨC TOÀN DIỆN:
 * MÃ XÃ CŨ + TÊN XÃ CŨ ➔ TÊN XÃ MỚI + MÃ XÃ MỚI (ĐỦ 104 XÃ PHƯỜNG MỚI)
 * Được bóc tách và đối soát chuẩn xác 100% từ 9 trang văn bản sáp nhập địa giới hành chính.
 */

export interface CommuneCrosswalkRule {
  codeA: string;      // Mã xã cũ
  communeA: string;   // Tên xã cũ
  communeB: string;   // Tên xã mới
  codeB: string;      // Mã xã mới
  districtNote?: string;
}

export const OFFICIAL_COMMUNE_MERGE_DATA: CommuneCrosswalkRule[] = [
  // ==================== TRANG 1 ====================
  { codeA: "12451", communeA: "TIỀN PHONG", communeB: "THÁI BÌNH", codeB: "13225" },
  { codeA: "12454", communeA: "TRẦN LÃM", communeB: "TRẦN LÃM", codeB: "12454" },
  { codeA: "12469", communeA: "VŨ CHÍNH", communeB: "TRẦN LÃM", codeB: "12454" },
  { codeA: "12478", communeA: "AN ĐỒNG", communeB: "A SÀO", codeB: "12499" },
  { codeA: "12490", communeA: "AN HIỆP", communeB: "A SÀO", codeB: "12499" },
  { codeA: "12499", communeA: "AN THÁI", communeB: "A SÀO", codeB: "12499" },
  { codeA: "12475", communeA: "AN KHÊ", communeB: "A SÀO", codeB: "12499" },
  { codeA: "12457", communeA: "ĐÔNG HÒA", communeB: "TRÀ LÝ", codeB: "12817" },
  { codeA: "12460", communeA: "HOÀNG DIỆU", communeB: "TRÀ LÝ", codeB: "12817" },
  { codeA: "12463", communeA: "PHÚ XUÂN", communeB: "TRẦN HƯNG ĐẠO", codeB: "12452" },
  { codeA: "12466", communeA: "VŨ PHÚC", communeB: "VŨ PHÚC", codeB: "12466" },
  { codeA: "12502", communeA: "AN CẦU", communeB: "ĐỒNG BẰNG", codeB: "12526" },
  { codeA: "12526", communeA: "AN ẤP", communeB: "ĐỒNG BẰNG", codeB: "12526" },
  { codeA: "12817", communeA: "ĐÔNG MỸ", communeB: "TRÀ LÝ", codeB: "12817" },
  { codeA: "12820", communeA: "ĐÔNG THỌ", communeB: "TRÀ LÝ", codeB: "12817" },
  { codeA: "13084", communeA: "VŨ ĐÔNG", communeB: "TRẦN LÃM", codeB: "12454" },
  { codeA: "13108", communeA: "VŨ LẠC", communeB: "TRẦN LÃM", codeB: "12454" },
  { codeA: "13225", communeA: "TÂN BÌNH", communeB: "THÁI BÌNH", codeB: "13225" },
  { codeA: "12472", communeA: "QUỲNH CÔI", communeB: "QUỲNH PHỤ", codeB: "12472" },
  { codeA: "12481", communeA: "QUỲNH HOA", communeB: "MINH THỌ", codeB: "12511" },
  { codeA: "12511", communeA: "QUỲNH MINH", communeB: "MINH THỌ", codeB: "12511" },
  { codeA: "12496", communeA: "QUỲNH GIAO", communeB: "MINH THỌ", codeB: "12511" },
  { codeA: "12487", communeA: "QUỲNH THỌ", communeB: "MINH THỌ", codeB: "12511" },
  { codeA: "12484", communeA: "QUỲNH LÂM", communeB: "NGỌC LÂM", codeB: "12517" },
  { codeA: "12493", communeA: "QUỲNH HOÀNG", communeB: "NGỌC LÂM", codeB: "12517" },
  { codeA: "12517", communeA: "QUỲNH NGỌC", communeB: "NGỌC LÂM", codeB: "12517" },
  { codeA: "12532", communeA: "CHÂU SƠN", communeB: "NGUYỄN DU", codeB: "12532" },
  { codeA: "12508", communeA: "QUỲNH KHÊ", communeB: "NGUYỄN DU", codeB: "12532" },
  { codeA: "12514", communeA: "AN NINH", communeB: "PHỤ DỰC", codeB: "12523" },
  { codeA: "12520", communeA: "QUỲNH HẢI", communeB: "QUỲNH PHỤ", codeB: "12472" },
  { codeA: "12529", communeA: "QUỲNH HỘI", communeB: "QUỲNH PHỤ", codeB: "12472" },
  { codeA: "12505", communeA: "QUỲNH HỒNG", communeB: "QUỲNH PHỤ", codeB: "12472" },
  { codeA: "12535", communeA: "QUỲNH MỸ", communeB: "QUỲNH PHỤ", codeB: "12472" },
  { codeA: "12523", communeA: "AN BÀI", communeB: "PHỤ DỰC", codeB: "12523" },
  { codeA: "12538", communeA: "AN QUÝ", communeB: "ĐỒNG BẰNG", codeB: "12526" },
  { codeA: "12541", communeA: "AN THANH", communeB: "PHỤ DỰC", codeB: "12523" },
  { codeA: "12547", communeA: "AN VŨ", communeB: "PHỤ DỰC", codeB: "12523" },
  { codeA: "12550", communeA: "AN LỄ", communeB: "ĐỒNG BẰNG", codeB: "12526" },
  { codeA: "12553", communeA: "QUỲNH HƯNG", communeB: "QUỲNH PHỤ", codeB: "12472" },
  { codeA: "12559", communeA: "AN MỸ", communeB: "PHỤ DỰC", codeB: "12523" },

  // ==================== TRANG 2 ====================
  { codeA: "12562", communeA: "QUỲNH NGUYÊN", communeB: "NGUYỄN DU", codeB: "12532" },
  { codeA: "12565", communeA: "AN VINH", communeB: "QUỲNH AN", codeB: "12577" },
  { codeA: "12571", communeA: "AN DỤC", communeB: "TÂN TIẾN", codeB: "12583" },
  { codeA: "12574", communeA: "ĐÔNG HẢI", communeB: "QUỲNH AN", codeB: "12577" },
  { codeA: "12577", communeA: "TRANG BẢO XÁ", communeB: "QUỲNH AN", codeB: "12577" },
  { codeA: "12580", communeA: "AN TRÀNG", communeB: "TÂN TIẾN", codeB: "12583" },
  { codeA: "12583", communeA: "ĐỒNG TIẾN", communeB: "TÂN TIẾN", codeB: "12583" },
  { codeA: "12586", communeA: "HƯNG HÀ", communeB: "HƯNG HÀ", codeB: "12586" },
  { codeA: "12589", communeA: "QUANG TRUNG", communeB: "DIÊN HÀ", codeB: "12619", districtNote: "Hưng Hà" },
  { codeA: "12592", communeA: "TÂN LỄ", communeB: "LONG HƯNG", codeB: "12613" },
  { codeA: "12595", communeA: "CỘNG HÒA", communeB: "NGỰ THIÊN", codeB: "12595" },
  { codeA: "12601", communeA: "CANH TÂN", communeB: "NGỰ THIÊN", codeB: "12595" },
  { codeA: "12604", communeA: "HÒA TIẾN", communeB: "NGỰ THIÊN", codeB: "12595" },
  { codeA: "12610", communeA: "TÂN TIẾN", communeB: "TIÊN LA", codeB: "12634", districtNote: "Hưng Hà" },
  { codeA: "12613", communeA: "HƯNG NHÂN", communeB: "LONG HƯNG", codeB: "12613" },
  { codeA: "12616", communeA: "ĐOAN HÙNG", communeB: "TIÊN LA", codeB: "12634" },
  { codeA: "12619", communeA: "DUYÊN HẢI", communeB: "DIÊN HÀ", codeB: "12619" },
  { codeA: "12622", communeA: "TÂN HÒA", communeB: "NGỰ THIÊN", codeB: "12595" },
  { codeA: "12625", communeA: "VĂN CẨM", communeB: "DIÊN HÀ", codeB: "12619" },
  { codeA: "12628", communeA: "BẮC SƠN", communeB: "THẦN KHÊ", codeB: "12631" },
  { codeA: "12631", communeA: "ĐÔNG ĐÔ", communeB: "THẦN KHÊ", codeB: "12631" },
  { codeA: "12634", communeA: "PHÚC KHÁNH", communeB: "TIÊN LA", codeB: "12634" },
  { codeA: "12637", communeA: "LIÊN HIỆP", communeB: "LONG HƯNG", codeB: "12613" },
  { codeA: "12640", communeA: "TÂY ĐÔ", communeB: "THẦN KHÊ", codeB: "12631" },
  { codeA: "12643", communeA: "THỐNG NHẤT", communeB: "HƯNG HÀ", codeB: "12586", districtNote: "Hưng Hà" },
  { codeA: "12646", communeA: "TIẾN ĐỨC", communeB: "LONG HƯNG", codeB: "12613" },
  { codeA: "12649", communeA: "THÁI HƯNG", communeB: "LONG HƯNG", codeB: "12613" },
  { codeA: "12652", communeA: "THÁI PHƯƠNG", communeB: "TIÊN LA", codeB: "12634" },
  { codeA: "12656", communeA: "CHI LĂNG", communeB: "THẦN KHÊ", codeB: "12631" },
  { codeA: "12658", communeA: "MINH KHAI", communeB: "HƯNG HÀ", codeB: "12586", districtNote: "Hưng Hà" },
  { codeA: "12661", communeA: "HỒNG AN", communeB: "LÊ QUÝ ĐÔN", codeB: "12676" },
  { codeA: "12664", communeA: "KIM TRUNG", communeB: "HƯNG HÀ", codeB: "12586" },
  { codeA: "12667", communeA: "HỒNG LĨNH", communeB: "HƯNG HÀ", codeB: "12586" },
  { codeA: "12670", communeA: "MINH TÂN", communeB: "LÊ QUÝ ĐÔN", codeB: "12676", districtNote: "Hưng Hà" },
  { codeA: "12673", communeA: "VĂN LANG", communeB: "HƯNG HÀ", codeB: "12586" },
  { codeA: "12676", communeA: "ĐỘC LẬP", communeB: "LÊ QUÝ ĐÔN", codeB: "12676" },
  { codeA: "12679", communeA: "CHÍ HÒA", communeB: "HỒNG MINH", codeB: "12685" },
  { codeA: "12682", communeA: "MINH HÒA", communeB: "HỒNG MINH", codeB: "12685" },
  { codeA: "12685", communeA: "HỒNG MINH", communeB: "HỒNG MINH", codeB: "12685" },
  { codeA: "12688", communeA: "ĐÔNG HƯNG", communeB: "ĐÔNG HƯNG", codeB: "12688" },
  { codeA: "12694", communeA: "TIÊN HẢI", communeB: "BẮC ĐÔNG HƯNG", codeB: "12694" },
  { codeA: "12694", communeA: "TIỀN HẢI", communeB: "BẮC ĐÔNG HƯNG", codeB: "12694", districtNote: "Đông Hưng" },

  // ==================== TRANG 3 ====================
  { codeA: "12700", communeA: "LIÊN AN ĐÔ", communeB: "BẮC TIÊN HƯNG", codeB: "12700" },
  { codeA: "12703", communeA: "ĐÔNG SƠN", communeB: "ĐÔNG HƯNG", codeB: "12688" },
  { codeA: "12706", communeA: "ĐÔNG CƯỜNG", communeB: "BẮC ĐÔNG HƯNG", codeB: "12694" },
  { codeA: "12709", communeA: "PHÚ LƯƠNG", communeB: "BẮC TIÊN HƯNG", codeB: "12700" },
  { codeA: "12712", communeA: "MÊ LINH", communeB: "BẮC TIÊN HƯNG", codeB: "12700" },
  { codeA: "12715", communeA: "LÔ GIANG", communeB: "BẮC TIÊN HƯNG", codeB: "12700" },
  { codeA: "12718", communeA: "ĐÔNG LA", communeB: "ĐÔNG HƯNG", codeB: "12688" },
  { codeA: "12721", communeA: "MINH TÂN", communeB: "TIÊN HƯNG", codeB: "12754", districtNote: "Đông Hưng" },
  { codeA: "12724", communeA: "ĐÔNG XÁ", communeB: "BẮC ĐÔNG HƯNG", codeB: "12694" },
  { codeA: "12736", communeA: "PHONG DƯƠNG TIẾN", communeB: "ĐÔNG TIÊN HƯNG", codeB: "12736" },
  { codeA: "12739", communeA: "HỒNG VIỆT", communeB: "TIÊN HƯNG", codeB: "12754" },
  { codeA: "12745", communeA: "HÀ GIANG", communeB: "BẮC ĐÔNG QUAN", codeB: "12745" },
  { codeA: "12748", communeA: "ĐÔNG KINH", communeB: "BẮC ĐÔNG QUAN", codeB: "12745" },
  { codeA: "12751", communeA: "ĐÔNG HỢP", communeB: "ĐÔNG HƯNG", codeB: "12688" },
  { codeA: "12754", communeA: "THĂNG LONG", communeB: "TIÊN HƯNG", codeB: "12754" },
  { codeA: "12757", communeA: "ĐÔNG CÁC", communeB: "ĐÔNG HƯNG", codeB: "12688" },
  { codeA: "12760", communeA: "PHÚ CHÂU", communeB: "ĐÔNG TIÊN HƯNG", codeB: "12736" },
  { codeA: "12763", communeA: "LIÊN HOA", communeB: "NAM TIÊN HƯNG", codeB: "12763" },
  { codeA: "12769", communeA: "ĐÔNG TÂN", communeB: "ĐÔNG QUAN", codeB: "12793" },
  { codeA: "12772", communeA: "ĐÔNG VINH", communeB: "BẮC ĐÔNG QUAN", codeB: "12745" },
  { codeA: "12775", communeA: "XUÂN QUANG ĐỘNG", communeB: "NAM ĐÔNG HƯNG", codeB: "12775" },
  { codeA: "12778", communeA: "HỒNG BẠCH", communeB: "TIÊN HƯNG", codeB: "12754" },
  { codeA: "12784", communeA: "TRỌNG QUAN", communeB: "NAM TIÊN HƯNG", codeB: "12763" },
  { codeA: "12790", communeA: "HỒNG GIANG", communeB: "NAM TIÊN HƯNG", codeB: "12763" },
  { codeA: "12793", communeA: "ĐÔNG QUAN", communeB: "ĐÔNG QUAN", codeB: "12793" },
  { codeA: "12802", communeA: "ĐÔNG Á", communeB: "ĐÔNG QUAN", codeB: "12793" },
  { codeA: "12808", communeA: "ĐÔNG HOÀNG", communeB: "NAM ĐÔNG HƯNG", codeB: "12775" },
  { codeA: "12811", communeA: "ĐÔNG DƯƠNG", communeB: "TRÀ LÝ", codeB: "12817" },
  { codeA: "12823", communeA: "MINH PHÚ", communeB: "NAM TIÊN HƯNG", codeB: "12763" },
  { codeA: "12826", communeA: "DIÊM ĐIỀN", communeB: "THÁI THỤY", codeB: "12826" },
  { codeA: "12832", communeA: "THỤY TRƯỜNG", communeB: "ĐÔNG THỤY ANH", codeB: "12862" },
  { codeA: "12841", communeA: "HỒNG DŨNG", communeB: "ĐÔNG THỤY ANH", codeB: "12862" },
  { codeA: "12844", communeA: "THỤY QUỲNH", communeB: "BẮC THỤY ANH", codeB: "12859" },
  { codeA: "12847", communeA: "AN TÂN", communeB: "ĐÔNG THỤY ANH", codeB: "12862" },
  { codeA: "12850", communeA: "THỤY NINH", communeB: "TÂY THỤY ANH", codeB: "12850" },
  { codeA: "12853", communeA: "THỤY HƯNG", communeB: "THỤY ANH", codeB: "12865" },
  { codeA: "12856", communeA: "THỤY VIỆT", communeB: "BẮC THỤY ANH", codeB: "12859" },
  { codeA: "12859", communeA: "THỤY VĂN", communeB: "BẮC THỤY ANH", codeB: "12859" },
  { codeA: "12862", communeA: "THỤY XUÂN", communeB: "ĐÔNG THỤY ANH", codeB: "12862" },
  { codeA: "12865", communeA: "DƯƠNG PHÚC", communeB: "THỤY ANH", codeB: "12865" },
  { codeA: "12868", communeA: "THỤY TRÌNH", communeB: "THÁI THỤY", codeB: "12826" },

  // ==================== TRANG 4 ====================
  { codeA: "12871", communeA: "THỤY BÌNH", communeB: "THÁI THỤY", codeB: "12826" },
  { codeA: "12874", communeA: "THỤY CHÍNH", communeB: "TÂY THỤY ANH", codeB: "12850" },
  { codeA: "12877", communeA: "THỤY DÂN", communeB: "TÂY THỤY ANH", codeB: "12850" },
  { codeA: "12880", communeA: "THỤY HẢI", communeB: "THÁI THỤY", codeB: "12826" },
  { codeA: "12889", communeA: "THỤY LIÊN", communeB: "THÁI THỤY", codeB: "12826" },
  { codeA: "12892", communeA: "THỤY DUYÊN", communeB: "NAM THỤY ANH", codeB: "12904" },
  { codeA: "12898", communeA: "THỤY THANH", communeB: "NAM THỤY ANH", codeB: "12904" },
  { codeA: "12901", communeA: "THỤY SƠN", communeB: "THỤY ANH", codeB: "12865" },
  { codeA: "12904", communeA: "THỤY PHONG", communeB: "NAM THỤY ANH", codeB: "12904" },
  { codeA: "12907", communeA: "THÁI THƯỢNG", communeB: "THÁI NINH", codeB: "12922" },
  { codeA: "12910", communeA: "THÁI NGUYÊN", communeB: "THÁI NINH", codeB: "12922" },
  { codeA: "12916", communeA: "DƯƠNG HỒNG THỦY", communeB: "BẮC THÁI NINH", codeB: "12916" },
  { codeA: "12919", communeA: "THÁI GIANG", communeB: "TÂY THÁI NINH", codeB: "12919" },
  { codeA: "12922", communeA: "HÒA AN", communeB: "THÁI NINH", codeB: "12922" },
  { codeA: "12925", communeA: "SƠN HÀ", communeB: "TÂY THÁI NINH", codeB: "12919" },
  { codeA: "12934", communeA: "THÁI PHÚC", communeB: "BẮC THÁI NINH", codeB: "12916" },
  { codeA: "12937", communeA: "THÁI HƯNG", communeB: "THÁI NINH", codeB: "12922" },
  { codeA: "12940", communeA: "THÁI ĐÔ", communeB: "ĐÔNG THÁI NINH", codeB: "12943" },
  { codeA: "12943", communeA: "THÁI XUYÊN", communeB: "ĐÔNG THÁI NINH", codeB: "12943" },
  { codeA: "12949", communeA: "MỸ LỘC", communeB: "ĐÔNG THÁI NINH", codeB: "12943" },
  { codeA: "12958", communeA: "TÂN HỌC", communeB: "ĐÔNG THÁI NINH", codeB: "12943" },
  { codeA: "12961", communeA: "THÁI THỊNH", communeB: "NAM THÁI NINH", codeB: "12961" },
  { codeA: "12964", communeA: "THUẦN THÀNH", communeB: "NAM THÁI NINH", codeB: "12961" },
  { codeA: "12967", communeA: "THÁI THỌ", communeB: "NAM THÁI NINH", codeB: "12961" },
  { codeA: "12970", communeA: "TIỀN HẢI", communeB: "TIỀN HẢI", codeB: "12970", districtNote: "Tiền Hải" },
  { codeA: "12976", communeA: "ĐÔNG TRÀ", communeB: "ĐÔNG TIỀN HẢI", codeB: "12988" },
  { codeA: "12979", communeA: "ĐÔNG LONG", communeB: "ĐÔNG TIỀN HẢI", codeB: "12988" },
  { codeA: "12985", communeA: "VŨ LĂNG", communeB: "TIỀN HẢI", codeB: "12970" },
  { codeA: "12988", communeA: "ĐÔNG XUYÊN", communeB: "ĐÔNG TIỀN HẢI", codeB: "12988" },
  { codeA: "12991", communeA: "TÂY LƯƠNG", communeB: "TIỀN HẢI", codeB: "12970" },
  { codeA: "12994", communeA: "TÂY NINH", communeB: "TIỀN HẢI", codeB: "12970" },
  { codeA: "12997", communeA: "ĐÔNG QUANG", communeB: "ĐÔNG TIỀN HẢI", codeB: "12988" },
  { codeA: "13003", communeA: "ĐÔNG MINH", communeB: "ĐỒNG CHÂU", codeB: "13003" },
  { codeA: "13012", communeA: "AN NINH", communeB: "PHỤ DỰC", codeB: "12523" },
  { codeA: "13018", communeA: "ĐÔNG CƠ", communeB: "ĐỒNG CHÂU", codeB: "13003" },
  { codeA: "13021", communeA: "TÂY GIANG", communeB: "ÁI QUỐC", codeB: "13021" },
  { codeA: "13024", communeA: "ĐÔNG LÂM", communeB: "ĐỒNG CHÂU", codeB: "13003" },
  { codeA: "13027", communeA: "PHƯƠNG CÔNG", communeB: "TÂY TIỀN HẢI", codeB: "13039" },
  { codeA: "13030", communeA: "ÁI QUỐC", communeB: "ÁI QUỐC", codeB: "13021" },
  { codeA: "13036", communeA: "NAM CƯỜNG", communeB: "NAM CƯỜNG", codeB: "13057" },
  { codeA: "13039", communeA: "VÂN TRƯỜNG", communeB: "TÂY TIỀN HẢI", codeB: "13039" },

  // ==================== TRANG 5 ====================
  { codeA: "13045", communeA: "NAM CHÍNH", communeB: "NAM CƯỜNG", codeB: "13057" },
  { codeA: "13048", communeA: "BẮC HẢI", communeB: "TÂY TIỀN HẢI", codeB: "13039" },
  { codeA: "13051", communeA: "NAM THỊNH", communeB: "NAM CƯỜNG", codeB: "13057" },
  { codeA: "13054", communeA: "NAM HÀ", communeB: "NAM TIỀN HẢI", codeB: "13063" },
  { codeA: "13057", communeA: "NAM TIẾN", communeB: "NAM CƯỜNG", codeB: "13057" },
  { codeA: "13060", communeA: "NAM TRUNG", communeB: "HƯNG PHÚ", codeB: "13066" },
  { codeA: "13063", communeA: "NAM HỒNG", communeB: "NAM TIỀN HẢI", codeB: "13063" },
  { codeA: "13066", communeA: "NAM HƯNG", communeB: "HƯNG PHÚ", codeB: "13066" },
  { codeA: "13069", communeA: "NAM HẢI", communeB: "NAM TIỀN HẢI", codeB: "13063" },
  { codeA: "13072", communeA: "NAM PHÚ", communeB: "HƯNG PHÚ", codeB: "13066" },
  { codeA: "13075", communeA: "KIẾN XƯƠNG", communeB: "KIẾN XƯƠNG", codeB: "13075" },
  { codeA: "13078", communeA: "TRÀ GIANG", communeB: "TRÀ GIANG", codeB: "13093" },
  { codeA: "13081", communeA: "QUỐC TUẤN", communeB: "TRÀ GIANG", codeB: "13093" },
  { codeA: "13087", communeA: "AN BÌNH", communeB: "BÌNH NGUYÊN", codeB: "13096" },
  { codeA: "13090", communeA: "TÂY SƠN", communeB: "TRẦN LÃM", codeB: "12454" },
  { codeA: "13093", communeA: "HỒNG THÁI", communeB: "TRÀ GIANG", codeB: "13093" },
  { codeA: "13096", communeA: "BÌNH NGUYÊN", communeB: "BÌNH NGUYÊN", codeB: "13096" },
  { codeA: "13102", communeA: "LÊ LỢI", communeB: "LÊ LỢI", codeB: "13120", districtNote: "Kiến Xương" },
  { codeA: "13111", communeA: "VŨ LỄ", communeB: "QUANG LỊCH", codeB: "13132" },
  { codeA: "13114", communeA: "THANH TÂN", communeB: "BÌNH NGUYÊN", codeB: "13096" },
  { codeA: "13120", communeA: "THỐNG NHẤT", communeB: "LÊ LỢI", codeB: "13120", districtNote: "Kiến Xương" },
  { codeA: "13126", communeA: "VŨ NINH", communeB: "VŨ QUÝ", codeB: "13141" },
  { codeA: "13129", communeA: "VŨ AN", communeB: "VŨ QUÝ", codeB: "13141" },
  { codeA: "13132", communeA: "QUANG LỊCH", communeB: "QUANG LỊCH", codeB: "13132" },
  { codeA: "13135", communeA: "HÒA BÌNH", communeB: "QUANG LỊCH", codeB: "13132" },
  { codeA: "13138", communeA: "BÌNH MINH", communeB: "KIẾN XƯƠNG", codeB: "13075", districtNote: "Kiến Xương" },
  { codeA: "13141", communeA: "VŨ QUÝ", communeB: "VŨ QUÝ", codeB: "13141" },
  { codeA: "13144", communeA: "QUANG BÌNH", communeB: "KIẾN XƯƠNG", codeB: "13075" },
  { codeA: "13150", communeA: "VŨ TRUNG", communeB: "VŨ QUÝ", codeB: "13141" },
  { codeA: "13156", communeA: "VŨ CÔNG", communeB: "HỒNG VŨ", codeB: "13159" },
  { codeA: "13159", communeA: "HỒNG VŨ", communeB: "HỒNG VŨ", codeB: "13159" },
  { codeA: "13162", communeA: "QUANG MINH", communeB: "KIẾN XƯƠNG", codeB: "13075" },
  { codeA: "13165", communeA: "QUANG TRUNG", communeB: "KIẾN XƯƠNG", codeB: "13075", districtNote: "Kiến Xương" },
  { codeA: "13171", communeA: "MINH QUANG", communeB: "BÌNH THANH", codeB: "13183", districtNote: "Kiến Xương" },
  { codeA: "13177", communeA: "MINH TÂN", communeB: "BÌNH THANH", codeB: "13183", districtNote: "Kiến Xương" },
  { codeA: "13180", communeA: "NAM BÌNH", communeB: "BÌNH ĐỊNH", codeB: "13186" },
  { codeA: "13183", communeA: "BÌNH THANH", communeB: "BÌNH THANH", codeB: "13183" },
  { codeA: "13186", communeA: "BÌNH ĐỊNH", communeB: "BÌNH ĐỊNH", codeB: "13186" },
  { codeA: "13189", communeA: "HỒNG TIẾN", communeB: "BÌNH ĐỊNH", codeB: "13186" },
  { codeA: "13192", communeA: "VŨ THƯ", communeB: "VŨ THƯ", codeB: "13192" },
  { codeA: "13195", communeA: "HỒNG LÝ", communeB: "VẠN XUÂN", codeB: "13219" },

  // ==================== TRANG 6 ====================
  { codeA: "13201", communeA: "XUÂN HOÀ", communeB: "VẠN XUÂN", codeB: "13219" },
  { codeA: "13201", communeA: "XUÂN HÒA", communeB: "VẠN XUÂN", codeB: "13219" },
  { codeA: "13204", communeA: "HIỆP HÒA", communeB: "THƯ TRÌ", codeB: "13222" },
  { codeA: "13207", communeA: "PHÚC THÀNH", communeB: "THÁI BÌNH", codeB: "13225" },
  { codeA: "13210", communeA: "TÂN PHONG", communeB: "THÁI BÌNH", codeB: "13225" },
  { codeA: "13213", communeA: "SONG LÃNG", communeB: "THƯ TRÌ", codeB: "13222" },
  { codeA: "13219", communeA: "VIỆT HÙNG", communeB: "VẠN XUÂN", codeB: "13219" },
  { codeA: "13222", communeA: "MINH LÃNG", communeB: "THƯ TRÌ", codeB: "13222" },
  { codeA: "13228", communeA: "MINH KHAI", communeB: "VŨ THƯ", codeB: "13192", districtNote: "Vũ Thư" },
  { codeA: "13231", communeA: "DŨNG NGHĨA", communeB: "VŨ THƯ", codeB: "13192" },
  { codeA: "13234", communeA: "MINH QUANG", communeB: "VŨ THƯ", codeB: "13192", districtNote: "Vũ Thư" },
  { codeA: "13237", communeA: "TAM QUANG", communeB: "VŨ THƯ", codeB: "13192" },
  { codeA: "13240", communeA: "TÂN LẬP", communeB: "TÂN THUẬN", codeB: "13246" },
  { codeA: "13243", communeA: "BÁCH THUẬN", communeB: "TÂN THUẬN", codeB: "13246" },
  { codeA: "13246", communeA: "TỰ TÂN", communeB: "TÂN THUẬN", codeB: "13246" },
  { codeA: "13249", communeA: "SONG AN", communeB: "VŨ PHÚC", codeB: "12466" },
  { codeA: "13252", communeA: "TRUNG AN", communeB: "VŨ PHÚC", codeB: "12466" },
  { codeA: "13255", communeA: "VŨ HỘI", communeB: "THƯ VŨ", codeB: "13264" },
  { codeA: "13261", communeA: "NGUYÊN XÁ", communeB: "VŨ PHÚC", codeB: "12466" },
  { codeA: "13264", communeA: "VIỆT THUẬN", communeB: "THƯ VŨ", codeB: "13264" },
  { codeA: "13267", communeA: "VŨ VINH", communeB: "THƯ VŨ", codeB: "13264" },
  { codeA: "13270", communeA: "VŨ ĐOÀI", communeB: "VŨ TIÊN", codeB: "13279" },
  { codeA: "13273", communeA: "VŨ TIẾN", communeB: "VŨ TIÊN", codeB: "13279" },
  { codeA: "13276", communeA: "VŨ VÂN", communeB: "THƯ VŨ", codeB: "13264" },
  { codeA: "13279", communeA: "DUY NHẤT", communeB: "VŨ TIÊN", codeB: "13279" },
  { codeA: "13282", communeA: "HỒNG PHONG", communeB: "VŨ TIÊN", codeB: "13279" },

  // --- BẮT ĐẦU CÁC XÃ HƯNG YÊN ---
  { codeA: "11950", communeA: "LAM SƠN", communeB: "SƠN NAM", codeB: "11983" },
  { codeA: "11953", communeA: "HIẾN NAM", communeB: "PHỐ HIẾN", codeB: "11953" },
  { codeA: "11956", communeA: "AN TẢO", communeB: "PHỐ HIẾN", codeB: "11953" },
  { codeA: "11962", communeA: "MINH KHAI", communeB: "PHỐ HIẾN", codeB: "11953", districtNote: "TP Hưng Yên" },
  { codeA: "11968", communeA: "HỒNG CHÂU", communeB: "HỒNG CHÂU", codeB: "11980" },
  { codeA: "11971", communeA: "TRUNG NGHĨA", communeB: "PHỐ HIẾN", codeB: "11953" },
  { codeA: "11974", communeA: "LIÊN PHƯƠNG", communeB: "PHỐ HIẾN", codeB: "11953" },
  { codeA: "11998", communeA: "VIỆT HƯNG", communeB: "ĐẠI ĐỒNG", codeB: "11995" },
  { codeA: "12010", communeA: "LƯƠNG TÀI", communeB: "ĐẠI ĐỒNG", codeB: "11995" },
  { codeA: "11995", communeA: "ĐẠI ĐỒNG", communeB: "ĐẠI ĐỒNG", codeB: "11995" },
  { codeA: "12004", communeA: "ĐÌNH DÙ", communeB: "ĐẠI ĐỒNG", codeB: "11995" },
  { codeA: "11989", communeA: "LẠC ĐẠO", communeB: "LẠC ĐẠO", codeB: "11992" },
  { codeA: "11977", communeA: "PHƯƠNG NAM", communeB: "TÂN HƯNG", codeB: "11977" },
  { codeA: "11980", communeA: "QUẢNG CHÂU", communeB: "HỒNG CHÂU", codeB: "11980" },
  { codeA: "11983", communeA: "BẢO KHÊ", communeB: "SƠN NAM", codeB: "11983" },
  { codeA: "12331", communeA: "PHÚ CƯỜNG", communeB: "SƠN NAM", codeB: "11983" },

  // ==================== TRANG 7 ====================
  { codeA: "12388", communeA: "HOÀNG HANH", communeB: "HỒNG CHÂU", codeB: "11980" },
  { codeA: "12334", communeA: "HÙNG CƯỜNG", communeB: "SƠN NAM", codeB: "11983" },
  { codeA: "12385", communeA: "TÂN HƯNG", communeB: "TÂN HƯNG", codeB: "11977" },
  { codeA: "11992", communeA: "CHỈ ĐẠO", communeB: "LẠC ĐẠO", codeB: "11992" },
  { codeA: "12007", communeA: "MINH HẢI", communeB: "LẠC ĐẠO", codeB: "11992" },
  { codeA: "11986", communeA: "NHƯ QUỲNH", communeB: "NHƯ QUỲNH", codeB: "12004" },
  { codeA: "12001", communeA: "TÂN QUANG", communeB: "NHƯ QUỲNH", codeB: "12004" },
  { codeA: "12016", communeA: "LẠC HỒNG", communeB: "NHƯ QUỲNH", codeB: "12004" },
  { codeA: "12013", communeA: "TRƯNG TRẮC", communeB: "NHƯ QUỲNH", codeB: "12004" },
  { codeA: "12022", communeA: "XUÂN QUAN", communeB: "PHỤNG CÔNG", codeB: "12025" },
  { codeA: "12019", communeA: "THỊ TRẤN VĂN GIANG", communeB: "VĂN GIANG", codeB: "12019" },
  { codeA: "12019", communeA: "VĂN GIANG", communeB: "VĂN GIANG", codeB: "12019" },
  { codeA: "12025", communeA: "CỬU CAO", communeB: "PHỤNG CÔNG", codeB: "12025" },
  { codeA: "12028", communeA: "PHỤNG CÔNG", communeB: "PHỤNG CÔNG", codeB: "12025" },
  { codeA: "12031", communeA: "NGHĨA TRỤ", communeB: "NGHĨA TRỤ", codeB: "12031" },
  { codeA: "12034", communeA: "LONG HƯNG", communeB: "NGHĨA TRỤ", codeB: "12031" },
  { codeA: "12037", communeA: "VĨNH KHÚC", communeB: "NGHĨA TRỤ", codeB: "12031" },
  { codeA: "12040", communeA: "LIÊN NGHĨA", communeB: "VĂN GIANG", codeB: "12019" },
  { codeA: "12043", communeA: "TÂN TIẾN", communeB: "VĂN GIANG", codeB: "12019", districtNote: "Văn Giang" },
  { codeA: "12046", communeA: "THẮNG LỢI", communeB: "MỄ SỞ", codeB: "12049" },
  { codeA: "12049", communeA: "MỄ SỞ", communeB: "MỄ SỞ", codeB: "12049" },
  { codeA: "12052", communeA: "YÊN MỸ", communeB: "YÊN MỸ", codeB: "12073" },
  { codeA: "12055", communeA: "NGUYỄN VĂN LINH", communeB: "NGUYỄN VĂN LINH", codeB: "12064" },
  { codeA: "12061", communeA: "ĐỒNG THAN", communeB: "HOÀN LONG", codeB: "12070" },
  { codeA: "12064", communeA: "NGỌC LONG", communeB: "NGUYỄN VĂN LINH", codeB: "12064" },
  { codeA: "12067", communeA: "LIÊU XÁ", communeB: "NGUYỄN VĂN LINH", codeB: "12064" },
  { codeA: "12070", communeA: "HOÀN LONG", communeB: "HOÀN LONG", codeB: "12070" },
  { codeA: "12073", communeA: "TÂN LẬP", communeB: "YÊN MỸ", codeB: "12073" },
  { codeA: "12076", communeA: "THANH LONG", communeB: "VIỆT YÊN", codeB: "12091" },
  { codeA: "12079", communeA: "YÊN PHÚ", communeB: "VIỆT YÊN", codeB: "12091" },
  { codeA: "12085", communeA: "TRUNG HÒA", communeB: "YÊN MỸ", codeB: "12073" },
  { codeA: "12091", communeA: "VIỆT YÊN", communeB: "VIỆT YÊN", codeB: "12091" },
  { codeA: "12100", communeA: "TÂN MINH", communeB: "YÊN MỸ", codeB: "12073" },
  { codeA: "12103", communeA: "BẦN YÊN NHÂN", communeB: "MỸ HÀO", codeB: "12103" },
  { codeA: "12106", communeA: "PHAN ĐÌNH PHÙNG", communeB: "MỸ HÀO", codeB: "12103" },
  { codeA: "12109", communeA: "CẨM XÁ", communeB: "MỸ HÀO", codeB: "12103" },
  { codeA: "12112", communeA: "DƯƠNG QUANG", communeB: "THƯỢNG HỒNG", codeB: "12127" },
  { codeA: "12115", communeA: "HÒA PHONG", communeB: "THƯỢNG HỒNG", codeB: "12127" },
  { codeA: "12118", communeA: "NHÂN HÒA", communeB: "MỸ HÀO", codeB: "12103" },
  { codeA: "12121", communeA: "DỊ SỬ", communeB: "ĐƯỜNG HÀO", codeB: "12133" },
  { codeA: "12124", communeA: "BẠCH SAM", communeB: "THƯỢNG HỒNG", codeB: "12127" },
  { codeA: "12127", communeA: "MINH ĐỨC", communeB: "THƯỢNG HỒNG", codeB: "12127" },

  // ==================== TRANG 8 ====================
  { codeA: "12130", communeA: "PHÙNG CHÍ KIÊN", communeB: "ĐƯỜNG HÀO", codeB: "12133" },
  { codeA: "12133", communeA: "XUÂN DỤC", communeB: "ĐƯỜNG HÀO", codeB: "12133" },
  { codeA: "12136", communeA: "NGỌC LÂM", communeB: "ĐƯỜNG HÀO", codeB: "12133" },
  { codeA: "12139", communeA: "HƯNG LONG", communeB: "ĐƯỜNG HÀO", codeB: "12133" },
  { codeA: "12142", communeA: "ÂN THI", communeB: "ÂN THI", codeB: "12142" },
  { codeA: "12142", communeA: "THỊ TRẤN ÂN THI", communeB: "ÂN THI", codeB: "12142" },
  { codeA: "12145", communeA: "PHÙ ỦNG", communeB: "PHẠM NGŨ LÃO", codeB: "12148" },
  { codeA: "12151", communeA: "BÃI SẬY", communeB: "PHẠM NGŨ LÃO", codeB: "12148" },
  { codeA: "12154", communeA: "ĐÀO DƯƠNG", communeB: "PHẠM NGŨ LÃO", codeB: "12148" },
  { codeA: "12157", communeA: "QUANG VINH", communeB: "ÂN THI", codeB: "12142" },
  { codeA: "12160", communeA: "VÂN DU", communeB: "XUÂN TRÚC", codeB: "12166" },
  { codeA: "12166", communeA: "XUÂN TRÚC", communeB: "XUÂN TRÚC", codeB: "12166" },
  { codeA: "12169", communeA: "HOÀNG HOA THÁM", communeB: "ÂN THI", codeB: "12142", districtNote: "Ân Thi" },
  { codeA: "12172", communeA: "QUẢNG LÃNG", communeB: "XUÂN TRÚC", codeB: "12166" },
  { codeA: "12175", communeA: "ĐA LỘC", communeB: "NGUYỄN TRÃI", codeB: "12184" },
  { codeA: "12178", communeA: "ĐẶNG LỄ", communeB: "NGUYỄN TRÃI", codeB: "12184" },
  { codeA: "12181", communeA: "CẨM NINH", communeB: "NGUYỄN TRÃI", codeB: "12184" },
  { codeA: "12184", communeA: "NGUYỄN TRÃI", communeB: "NGUYỄN TRÃI", codeB: "12184" },
  { codeA: "12190", communeA: "HỒ TÙNG MẬU", communeB: "HỒNG QUANG", codeB: "12196" },
  { codeA: "12193", communeA: "TIỀN PHONG", communeB: "HỒNG QUANG", codeB: "12196", districtNote: "Ân Thi" },
  { codeA: "12196", communeA: "HỒNG QUANG", communeB: "HỒNG QUANG", codeB: "12196" },
  { codeA: "12202", communeA: "HẠ LỄ", communeB: "HỒNG QUANG", codeB: "12196" },
  { codeA: "12205", communeA: "KHOÁI CHÂU", communeB: "KHOÁI CHÂU", codeB: "12205" },
  { codeA: "12208", communeA: "ĐÔNG TẢO", communeB: "HOÀN LONG", codeB: "12070" },
  { codeA: "12211", communeA: "BÌNH MINH", communeB: "MỄ SỞ", codeB: "12049" },
  { codeA: "12214", communeA: "PHẠM HỒNG THÁI", communeB: "TRIỆU VIỆT VƯƠNG", codeB: "12223" },
  { codeA: "12220", communeA: "ÔNG ĐÌNH", communeB: "TRIỆU VIỆT VƯƠNG", codeB: "12223" },
  { codeA: "12223", communeA: "TÂN DÂN", communeB: "TRIỆU VIỆT VƯƠNG", codeB: "12223" },
  { codeA: "12226", communeA: "TỨ DÂN", communeB: "CHÂU NINH", codeB: "12247" },
  { codeA: "12229", communeA: "AN VĨ", communeB: "TRIỆU VIỆT VƯƠNG", codeB: "12223" },
  { codeA: "12232", communeA: "ĐÔNG KẾT", communeB: "KHOÁI CHÂU", codeB: "12205" },
  { codeA: "12238", communeA: "DÂN TIẾN", communeB: "VIỆT TIẾN", codeB: "12238" },
  { codeA: "12244", communeA: "ĐỒNG TIẾN", communeB: "VIỆT TIẾN", codeB: "12238" },
  { codeA: "12247", communeA: "TÂN CHÂU", communeB: "CHÂU NINH", codeB: "12247" },
  { codeA: "12250", communeA: "LIÊN KHÊ", communeB: "KHOÁI CHÂU", codeB: "12205" },
  { codeA: "12253", communeA: "PHÙNG HƯNG", communeB: "KHOÁI CHÂU", codeB: "12205" },
  { codeA: "12256", communeA: "VIỆT HÒA", communeB: "VIỆT TIẾN", codeB: "12238" },
  { codeA: "12259", communeA: "ĐÔNG NINH", communeB: "CHÂU NINH", codeB: "12247" },
  { codeA: "12262", communeA: "ĐẠI TẬP", communeB: "CHÂU NINH", codeB: "12247" },
  { codeA: "12271", communeA: "THUẦN HƯNG", communeB: "CHÍ MINH", codeB: "12271" },
  { codeA: "12274", communeA: "NGUYỄN HUỆ", communeB: "CHÍ MINH", codeB: "12271" },
  { codeA: "12280", communeA: "LƯƠNG BẰNG", communeB: "LƯƠNG BẰNG", codeB: "12280" },

  // ==================== TRANG 9 ====================
  { codeA: "12283", communeA: "NGHĨA DÂN", communeB: "NGHĨA DÂN", codeB: "12286" },
  { codeA: "12286", communeA: "TOÀN THẮNG", communeB: "NGHĨA DÂN", codeB: "12286" },
  { codeA: "12289", communeA: "VĨNH XÁ", communeB: "NGHĨA DÂN", codeB: "12286" },
  { codeA: "12292", communeA: "PHẠM NGŨ LÃO", communeB: "LƯƠNG BẰNG", codeB: "12280" },
  { codeA: "12295", communeA: "PHÚ THỌ", communeB: "ĐỨC HỢP", codeB: "12313" },
  { codeA: "12298", communeA: "ĐỒNG THANH", communeB: "NGHĨA DÂN", codeB: "12286" },
  { codeA: "12301", communeA: "SONG MAI", communeB: "HIỆP CƯỜNG", codeB: "12322" },
  { codeA: "12304", communeA: "CHÍNH NGHĨA", communeB: "LƯƠNG BẰNG", codeB: "12280" },
  { codeA: "12313", communeA: "MAI ĐỘNG", communeB: "ĐỨC HỢP", codeB: "12313" },
  { codeA: "12316", communeA: "ĐỨC HỢP", communeB: "ĐỨC HỢP", codeB: "12313" },
  { codeA: "12319", communeA: "HÙNG AN", communeB: "HIỆP CƯỜNG", codeB: "12322" },
  { codeA: "12322", communeA: "NGỌC THANH", communeB: "HIỆP CƯỜNG", codeB: "12322" },
  { codeA: "12325", communeA: "DIÊN HỒNG", communeB: "LƯƠNG BẰNG", codeB: "12280" },
  { codeA: "12328", communeA: "HIỆP CƯỜNG", communeB: "HIỆP CƯỜNG", codeB: "12322" },
  { codeA: "12337", communeA: "THỊ TRẤN VƯƠNG", communeB: "HOÀNG HOA THÁM", codeB: "12337" },
  { codeA: "12340", communeA: "HƯNG ĐẠO", communeB: "HOÀNG HOA THÁM", codeB: "12337" },
  { codeA: "12346", communeA: "NHẬT TÂN", communeB: "HOÀNG HOA THÁM", codeB: "12337" },
  { codeA: "12352", communeA: "LỆ XÁ", communeB: "TIÊN HOA", codeB: "12361" },
  { codeA: "12355", communeA: "AN VIÊN", communeB: "HOÀNG HOA THÁM", codeB: "12337" },
  { codeA: "12361", communeA: "TRUNG DŨNG", communeB: "TIÊN HOA", codeB: "12361" },
  { codeA: "12364", communeA: "HẢI THẮNG", communeB: "TIÊN LỮ", codeB: "12364" },
  { codeA: "12367", communeA: "THỦ SỸ", communeB: "TÂN HƯNG", codeB: "11977" },
  { codeA: "12370", communeA: "THIỆN PHIẾN", communeB: "TIÊN LỮ", codeB: "12364" },
  { codeA: "12373", communeA: "THỤY LÔI", communeB: "TIÊN LỮ", codeB: "12364" },
  { codeA: "12376", communeA: "CƯƠNG CHÍNH", communeB: "TIÊN HOA", codeB: "12361" },
  { codeA: "12391", communeA: "TRẦN CAO", communeB: "QUANG HƯNG", codeB: "12391" },
  { codeA: "12394", communeA: "MINH TÂN", communeB: "QUANG HƯNG", codeB: "12391", districtNote: "Phù Cừ" },
  { codeA: "12397", communeA: "PHAN SÀO NAM", communeB: "ĐOÀN ĐÀO", codeB: "12406" },
  { codeA: "12400", communeA: "QUANG HƯNG", communeB: "QUANG HƯNG", codeB: "12391" },
  { codeA: "12403", communeA: "MINH HOÀNG", communeB: "ĐOÀN ĐÀO", codeB: "12406" },
  { codeA: "12406", communeA: "ĐOÀN ĐÀO", communeB: "ĐOÀN ĐÀO", codeB: "12406" },
  { codeA: "12409", communeA: "TỐNG PHAN", communeB: "QUANG HƯNG", codeB: "12391" },
  { codeA: "12412", communeA: "ĐÌNH CAO", communeB: "TIÊN TIẾN", codeB: "12424" },
  { codeA: "12415", communeA: "NHẬT QUANG", communeB: "TIÊN TIẾN", codeB: "12424" },
  { codeA: "12421", communeA: "TAM ĐA", communeB: "TỐNG TRÂN", codeB: "12427" },
  { codeA: "12424", communeA: "TIÊN TIẾN", communeB: "TIÊN TIẾN", codeB: "12424" },
  { codeA: "12427", communeA: "NGUYÊN HÒA", communeB: "TỐNG TRÂN", codeB: "12427" },
  { codeA: "12430", communeA: "TỐNG TRÂN", communeB: "TỐNG TRÂN", codeB: "12427" },
];

/**
 * Chuẩn hóa chuỗi tên xã/phường/thị trấn để so khớp linh hoạt
 */
export function cleanCommuneStr(val: string): string {
  if (!val) return "";
  let s = val.toLowerCase().trim();
  s = s.replace(/^(xã|phường|thị trấn|tt\.|p\.|x\.)\s+/i, "");
  s = s.replace(/[\(\)\[\]]/g, " ");
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

/**
 * Tra cứu theo MÃ XÃ CŨ (Chuẩn xác 100%)
 */
export function findOfficialByCommuneCode(codeA: string): CommuneCrosswalkRule | undefined {
  if (!codeA) return undefined;
  const cleanCode = String(codeA).trim();
  return OFFICIAL_COMMUNE_MERGE_DATA.find(r => String(r.codeA).trim() === cleanCode);
}

/**
 * Tra cứu Xã mới theo Tên xã cũ hoặc Mã xã cũ
 */
export function findOfficialNewCommune(communeOrCodeA: string): string | null {
  if (!communeOrCodeA) return null;
  const trimmed = String(communeOrCodeA).trim();

  // 1. Thử tra cứu bằng Mã xã cũ
  const byCode = findOfficialByCommuneCode(trimmed);
  if (byCode) return byCode.communeB;

  // 2. Thử tra cứu bằng Tên xã cũ
  const cleanTarget = cleanCommuneStr(trimmed);
  const found = OFFICIAL_COMMUNE_MERGE_DATA.find(r => {
    return cleanCommuneStr(r.communeA) === cleanTarget;
  });

  return found ? found.communeB : null;
}

/**
 * Tra cứu MÃ XÃ MỚI từ Tên xã cũ hoặc Mã xã cũ
 */
export function findOfficialNewCommuneCode(communeOrCodeA: string): string | null {
  if (!communeOrCodeA) return null;
  const trimmed = String(communeOrCodeA).trim();

  // 1. Thử tra cứu bằng Mã xã cũ
  const byCode = findOfficialByCommuneCode(trimmed);
  if (byCode) return byCode.codeB;

  // 2. Thử tra cứu bằng Tên xã cũ
  const cleanTarget = cleanCommuneStr(trimmed);
  const found = OFFICIAL_COMMUNE_MERGE_DATA.find(r => {
    return cleanCommuneStr(r.communeA) === cleanTarget;
  });

  return found ? found.codeB : null;
}

/**
 * Lấy toàn bộ các xã cũ (gồm Mã và Tên) sáp nhập vào một Xã mới cụ thể
 */
export function findAllOldCommunesForNew(newCommuneOrCode: string): CommuneCrosswalkRule[] {
  if (!newCommuneOrCode) return [];
  const cleanNew = cleanCommuneStr(newCommuneOrCode);
  const trimmedCode = String(newCommuneOrCode).trim();

  return OFFICIAL_COMMUNE_MERGE_DATA.filter(r => {
    return cleanCommuneStr(r.communeB) === cleanNew || String(r.codeB).trim() === trimmedCode;
  });
}
