// Danh mục Hệ thống ngành kinh tế Việt Nam (VSIC) theo Quyết định số 27/2018/QĐ-TTg

export const vsicRawData: Record<string, string> = {
  // CẤP 1 (A - U)
  "A": "Nông nghiệp, lâm nghiệp và thủy sản",
  "B": "Khai khoáng",
  "C": "Công nghiệp chế biến, chế tạo",
  "D": "Sản xuất và phân phối điện, khí đốt, nước nóng, hơi nước và điều hoà không khí",
  "E": "Cung cấp nước; hoạt động quản lý và xử lý rác thải, nước thải",
  "F": "Xây dựng",
  "G": "Bán buôn và bán lẻ; sửa chữa ô tô, mô tô, xe máy và xe có động cơ khác",
  "H": "Vận tải kho bãi",
  "I": "Dịch vụ lưu trú và ăn uống",
  "J": "Thông tin và truyền thông",
  "K": "Hoạt động tài chính, ngân hàng và bảo hiểm",
  "L": "Hoạt động kinh doanh bất động sản",
  "M": "Hoạt động chuyên môn, khoa học và công nghệ",
  "N": "Hoạt động hành chính và dịch vụ hỗ trợ",
  "O": "Hoạt động của Đảng, tổ chức chính trị - xã hội, quản lý nhà nước, an ninh quốc phòng",
  "P": "Giáo dục và đào tạo",
  "Q": "Y tế và hoạt động trợ giúp xã hội",
  "R": "Nghệ thuật, vui chơi và giải trí",
  "S": "Hoạt động dịch vụ khác",
  "T": "Hoạt động làm thuê các công việc trong các gia đình",
  "U": "Hoạt động của các tổ chức và cơ quan quốc tế",

  // CẤP 2 (01 - 99)
  "01": "Nông nghiệp và hoạt động dịch vụ có liên quan",
  "02": "Lâm nghiệp và hoạt động dịch vụ có liên quan",
  "03": "Khai thác, nuôi trồng thủy sản",
  "05": "Khai thác than cứng và than non",
  "06": "Khai thác dầu thô và khí đốt tự nhiên",
  "07": "Khai thác quặng kim loại",
  "08": "Khai khoáng khác",
  "09": "Hoạt động dịch vụ hỗ trợ khai khoáng",
  "10": "Sản xuất chế biến thực phẩm",
  "11": "Sản xuất đồ uống",
  "12": "Sản xuất sản phẩm thuốc lá",
  "13": "Dệt",
  "14": "Sản xuất trang phục",
  "15": "Sản xuất da và các sản phẩm có liên quan",
  "16": "Chế biến gỗ và sản xuất sản phẩm từ gỗ, tre, nứa",
  "17": "Sản xuất giấy và sản phẩm từ giấy",
  "18": "In, sao chép bản ghi các loại",
  "19": "Sản xuất than cốc, sản phẩm dầu mỏ tinh chế",
  "20": "Sản xuất hoá chất và sản phẩm hoá chất",
  "21": "Sản xuất thuốc, hoá dược và dược liệu",
  "22": "Sản xuất sản phẩm từ cao su và plastic",
  "23": "Sản xuất sản phẩm từ khoáng phi kim loại khác",
  "24": "Sản xuất kim loại",
  "25": "Sản xuất sản phẩm từ kim loại đúc sẵn",
  "26": "Sản xuất sản phẩm điện tử, máy vi tính và sản phẩm quang học",
  "27": "Sản xuất thiết bị điện",
  "28": "Sản xuất máy móc, thiết bị chưa được phân vào đâu",
  "29": "Sản xuất ô tô và xe có động cơ khác",
  "30": "Sản xuất phương tiện vận tải khác",
  "31": "Sản xuất giường, tủ, bàn, ghế",
  "32": "Công nghiệp chế biến, chế tạo khác",
  "33": "Sửa chữa, bảo dưỡng và lắp đặt máy móc, thiết bị",
  "35": "Sản xuất và phân phối điện, khí đốt, nước nóng, hơi nước",
  "36": "Khai thác, xử lý và cung cấp nước",
  "37": "Thoát nước và xử lý nước thải",
  "38": "Hoạt động thu gom, xử lý và tiêu huỷ rác thải; tái chế phế liệu",
  "39": "Xử lý ô nhiễm và hoạt động quản lý chất thải khác",
  "41": "Xây dựng nhà các loại",
  "42": "Xây dựng công trình kỹ thuật dân dụng",
  "43": "Hoạt động xây dựng chuyên dụng",
  "45": "Bán buôn, bán lẻ, sửa chữa ô tô, mô tô, xe máy và xe có động cơ khác",
  "46": "Bán buôn (trừ ô tô, mô tô, xe máy và xe có động cơ khác)",
  "47": "Bán lẻ (trừ ô tô, mô tô, xe máy và xe có động cơ khác)",
  "49": "Vận tải đường bộ và vận tải đường ống",
  "50": "Vận tải đường thuỷ",
  "51": "Vận tải hàng không",
  "52": "Kho bãi và các hoạt động hỗ trợ cho vận tải",
  "53": "Bưu chính và chuyển phát",
  "55": "Dịch vụ lưu trú",
  "56": "Dịch vụ ăn uống",
  "58": "Hoạt động xuất bản",
  "59": "Hoạt động điện ảnh, sản xuất chương trình truyền hình, ghi âm",
  "60": "Hoạt động phát thanh, truyền hình",
  "61": "Viễn thông",
  "62": "Lập trình máy vi tính, dịch vụ tư vấn và các hoạt động khác",
  "63": "Hoạt động dịch vụ thông tin",
  "64": "Hoạt động dịch vụ tài chính (trừ bảo hiểm và bảo hiểm xã hội)",
  "65": "Bảo hiểm, tái bảo hiểm và bảo hiểm xã hội",
  "66": "Hoạt động tài chính khác",
  "68": "Hoạt động kinh doanh bất động sản",
  "69": "Hoạt động pháp luật, kế toán và kiểm toán",
  "70": "Hoạt động của trụ sở sở hữu, hoạt động tư vấn quản lý",
  "71": "Hoạt động kiến trúc; kiểm tra và phân tích kỹ thuật",
  "72": "Nghiên cứu khoa học và phát triển công nghệ",
  "73": "Quảng cáo và nghiên cứu thị trường",
  "74": "Hoạt động chuyên môn, khoa học và công nghệ khác",
  "75": "Hoạt động thú y",
  "77": "Cho thuê máy móc, thiết bị và đồ dùng hữu hình khác",
  "78": "Hoạt động việc làm",
  "79": "Hoạt động của các đại lý du lịch, kinh doanh tua du lịch",
  "80": "Hoạt động điều tra và an ninh",
  "81": "Dịch vụ vệ sinh nhà cửa, công trình và cảnh quan",
  "82": "Hoạt động hành chính, hỗ trợ văn phòng và dịch vụ hỗ trợ kinh doanh khác",
  "84": "Hoạt động của Đảng, tổ chức chính trị - xã hội, quản lý nhà nước",
  "85": "Giáo dục và đào tạo",
  "86": "Hoạt động y tế",
  "87": "Hoạt động chăm sóc, điều dưỡng tập trung",
  "88": "Hoạt động trợ giúp xã hội không tập trung",
  "90": "Hoạt động sáng tác, nghệ thuật và giải trí",
  "91": "Hoạt động của thư viện, lưu trữ, bảo tàng",
  "92": "Hoạt động đánh bạc và cá cược",
  "93": "Hoạt động thể thao, vui chơi và giải trí",
  "94": "Hoạt động của các hiệp hội, tổ chức khác",
  "95": "Sửa chữa máy vi tính, đồ dùng cá nhân và gia đình",
  "96": "Hoạt động dịch vụ phục vụ cá nhân khác",
  "97": "Hoạt động làm thuê công việc gia đình trong các hộ gia đình",
  "98": "Hoạt động sản xuất sản phẩm tự tiêu dùng của hộ gia đình",
  "99": "Hoạt động của các tổ chức và cơ quan quốc tế",

  // CẤP 3, CẤP 4, CẤP 5 TIÊU BIỂU VÀ PHỔ BIẾN
  "101": "Chế biến, bảo quản thịt và các sản phẩm từ thịt",
  "1010": "Chế biến, bảo quản thịt và các sản phẩm từ thịt",
  "10101": "Chế biến, bảo quản thịt gia súc",
  "10102": "Chế biến, bảo quản thịt gia cầm",
  "107": "Sản xuất thực phẩm khác",
  "1071": "Sản xuất các loại bánh từ bột",
  "10710": "Sản xuất các loại bánh từ bột",
  "141": "May trang phục (trừ trang phục từ da lông thú)",
  "1410": "May trang phục (trừ trang phục từ da lông thú)",
  "14100": "May trang phục (trừ trang phục từ da lông thú)",
  "259": "Sản xuất sản phẩm khác bằng kim loại; các dịch vụ xử lý kim loại",
  "2592": "Gia công cơ khí; xử lý và tráng phủ kim loại",
  "25920": "Gia công cơ khí; xử lý và tráng phủ kim loại",
  "451": "Bán buôn, bán lẻ ô tô và xe có động cơ khác",
  "4511": "Bán buôn ô tô và xe có động cơ khác",
  "45110": "Bán buôn ô tô và xe có động cơ khác",
  "4512": "Bán lẻ ô tô con (loại 9 chỗ ngồi trở xuống)",
  "45120": "Bán lẻ ô tô con (loại 9 chỗ ngồi trở xuống)",
  "452": "Bảo dưỡng, sửa chữa ô tô và xe có động cơ khác",
  "4520": "Bảo dưỡng, sửa chữa ô tô và xe có động cơ khác",
  "45200": "Bảo dưỡng, sửa chữa ô tô và xe có động cơ khác",
  "454": "Bán, bảo dưỡng và sửa chữa mô tô, xe máy",
  "4540": "Bán, bảo dưỡng và sửa chữa mô tô, xe máy",
  "45401": "Bán buôn mô tô, xe máy",
  "45402": "Bán lẻ mô tô, xe máy",
  "45403": "Bảo dưỡng và sửa chữa mô tô, xe máy",
  "462": "Bán buôn nông, lâm sản nguyên liệu và động vật sống",
  "4620": "Bán buôn nông, lâm sản nguyên liệu và động vật sống",
  "46209": "Bán buôn thức ăn và nguyên liệu làm thức ăn cho gia súc, gia cầm, thủy sản",
  "471": "Bán lẻ trong các cửa hàng kinh doanh tổng hợp",
  "4711": "Bán lẻ trong siêu thị, trung tâm thương mại, cửa hàng tiện ích",
  "47110": "Bán lẻ lương thực, thực phẩm, đồ uống chiếm tỷ trọng lớn trong siêu thị mini",
  "472": "Bán lẻ thực phẩm, đồ uống, thuốc lá chiếm tỷ trọng lớn",
  "4721": "Bán lẻ lương thực trong các cửa hàng chuyên doanh",
  "47210": "Bán lẻ lương thực trong các cửa hàng chuyên doanh",
  "474": "Bán lẻ thiết bị công nghệ thông tin liên lạc",
  "4741": "Bán lẻ máy vi tính, thiết bị ngoại vi, phần mềm và thiết bị viễn thông",
  "47412": "Bán lẻ điện thoại di động và thiết bị viễn thông",
  "477": "Bán lẻ hàng hóa khác mới trong các cửa hàng chuyên doanh",
  "4772": "Bán lẻ thuốc, dụng cụ y tế, mỹ phẩm và vật phẩm vệ sinh",
  "47721": "Bán lẻ thuốc tân dược trong các cửa hàng chuyên doanh (nhà thuốc)",
  "493": "Vận tải đường bộ khác",
  "4933": "Vận tải hàng hoá bằng đường bộ",
  "49339": "Vận tải hàng hoá bằng xe tải khác",
  "561": "Nhà hàng và các dịch vụ ăn uống phục vụ lưu động",
  "5610": "Nhà hàng và các dịch vụ ăn uống phục vụ lưu động",
  "56101": "Quán ăn, nhà hàng, quán bún phở cơm phục vụ ăn uống tại chỗ",
  "56102": "Cửa hàng bán đồ ăn nhanh, phục vụ mang đi",
  "563": "Dịch vụ phục vụ đồ uống",
  "5630": "Dịch vụ phục vụ đồ uống",
  "56301": "Quán cà phê, giải khát",
  "56309": "Dịch vụ phục vụ đồ uống khác",
  "952": "Sửa chữa đồ dùng cá nhân và gia đình",
  "9521": "Sửa chữa thiết bị nghe nhìn điện tử gia đình",
  "95210": "Sửa chữa thiết bị nghe nhìn điện tử gia đình",
  "963": "Hoạt động dịch vụ phục vụ cá nhân khác",
  "9631": "Cắt tóc, làm đầu, gội đầu, làm móng và thẩm mỹ spa",
  "96310": "Cắt tóc, làm đầu, gội đầu, làm móng và thẩm mỹ spa"
};

// Bản đồ phân loại Ngành cấp 2 sang Ngành cấp 1 (A - U)
export const SECTOR_2_TO_1: Record<string, string> = {
  "01": "A", "02": "A", "03": "A",
  "05": "B", "06": "B", "07": "B", "08": "B", "09": "B",
  "10": "C", "11": "C", "12": "C", "13": "C", "14": "C", "15": "C", "16": "C", "17": "C", "18": "C", "19": "C",
  "20": "C", "21": "C", "22": "C", "23": "C", "24": "C", "25": "C", "26": "C", "27": "C", "28": "C", "29": "C",
  "30": "C", "31": "C", "32": "C", "33": "C",
  "35": "D",
  "36": "E", "37": "E", "38": "E", "39": "E",
  "41": "F", "42": "F", "43": "F",
  "45": "G", "46": "G", "47": "G",
  "49": "H", "50": "H", "51": "H", "52": "H", "53": "H",
  "55": "I", "56": "I",
  "58": "J", "59": "J", "60": "J", "61": "J", "62": "J", "63": "J",
  "64": "K", "65": "K", "66": "K",
  "68": "L",
  "69": "M", "70": "M", "71": "M", "72": "M", "73": "M", "74": "M", "75": "M",
  "77": "N", "78": "N", "79": "N", "80": "N", "81": "N", "82": "N",
  "84": "O",
  "85": "P",
  "86": "Q", "87": "Q", "88": "Q",
  "90": "R", "91": "R", "92": "R", "93": "R",
  "94": "S", "95": "S", "96": "S",
  "97": "T", "98": "T",
  "99": "U"
};

/**
 * Chuẩn hóa mã ngành: cắt khoảng trắng, loại bỏ dấu chấm/phẩy, chỉ giữ chữ hoặc số.
 */
export function normalizeSectorCode(raw: any): string {
  if (raw === null || raw === undefined) return "";
  let str = String(raw).trim().toUpperCase();
  // Nếu là chữ cái cấp 1 (A-U)
  if (/^[A-U]$/.test(str)) return str;
  // Loại bỏ các ký tự không phải chữ số
  str = str.replace(/[^\dA-Z]/g, "");
  return str;
}

/**
 * Xác định cấp của mã ngành (1 - 5).
 */
export function getSectorLevel(code: string): number {
  if (!code) return 0;
  if (/^[A-U]$/i.test(code)) return 1;
  const numLen = code.replace(/\D/g, "").length;
  if (numLen >= 1 && numLen <= 5) return numLen;
  return 0;
}

/**
 * Lấy mã ngành cấp 1 tương ứng (A - U).
 */
export function getParentSectorCode(code: string): string {
  if (!code) return "";
  const norm = normalizeSectorCode(code);
  if (/^[A-U]$/.test(norm)) return norm;

  const sec2 = norm.slice(0, 2);
  return SECTOR_2_TO_1[sec2] || "";
}

/**
 * Tra cứu tên ngành với cơ chế fallback từng cấp.
 */
export function lookupSectorNameWithFallback(code: string): { level: number; name: string; exactMatched?: boolean } {
  const norm = normalizeSectorCode(code);
  if (!norm) return { level: 0, name: "", exactMatched: false };

  // Tra cứu khớp chính xác
  if (vsicRawData[norm]) {
    return { level: getSectorLevel(norm), name: vsicRawData[norm], exactMatched: true };
  }

  // Fallback từ cấp 5 xuống cấp 4, 3, 2, 1
  if (norm.length >= 4 && vsicRawData[norm.slice(0, 4)]) {
    return { level: 4, name: vsicRawData[norm.slice(0, 4)], exactMatched: false };
  }
  if (norm.length >= 3 && vsicRawData[norm.slice(0, 3)]) {
    return { level: 3, name: vsicRawData[norm.slice(0, 3)], exactMatched: false };
  }
  if (norm.length >= 2 && vsicRawData[norm.slice(0, 2)]) {
    return { level: 2, name: vsicRawData[norm.slice(0, 2)], exactMatched: false };
  }

  const p1 = getParentSectorCode(norm);
  if (p1 && vsicRawData[p1]) {
    return { level: 1, name: vsicRawData[p1], exactMatched: false };
  }

  return { level: 0, name: "", exactMatched: false };
}

/**
 * Lấy phân cấp đầy đủ 5 cấp của một mã ngành.
 */
export function getSectorHierarchy(code: string): Record<string, { ma: string; ten: string }> {
  const norm = normalizeSectorCode(code);
  const result: Record<string, { ma: string; ten: string }> = {
    "1": { ma: "", ten: "" },
    "2": { ma: "", ten: "" },
    "3": { ma: "", ten: "" },
    "4": { ma: "", ten: "" },
    "5": { ma: "", ten: "" }
  };

  if (!norm) return result;

  const c1 = getParentSectorCode(norm);
  if (c1) result["1"] = { ma: c1, ten: vsicRawData[c1] || "" };

  if (norm.length >= 2) {
    const c2 = norm.slice(0, 2);
    result["2"] = { ma: c2, ten: vsicRawData[c2] || "" };
  }
  if (norm.length >= 3) {
    const c3 = norm.slice(0, 3);
    result["3"] = { ma: c3, ten: vsicRawData[c3] || "" };
  }
  if (norm.length >= 4) {
    const c4 = norm.slice(0, 4);
    result["4"] = { ma: c4, ten: vsicRawData[c4] || "" };
  }
  if (norm.length >= 5) {
    const c5 = norm.slice(0, 5);
    result["5"] = { ma: c5, ten: vsicRawData[c5] || vsicRawData[norm.slice(0, 4)] || "" };
  }

  return result;
}

/**
 * Gợi ý mã ngành dựa trên nội dung mô tả hoạt động thực tế.
 */
export function smartSuggestSectorByDescription(desc: string): { ma: string; ten: string; diem: number } | null {
  if (!desc || !desc.trim()) return null;
  const d = desc.toLowerCase().trim();

  // Danh mục từ khóa phổ biến
  const keywordRules: { keywords: string[]; ma: string; ten: string; diem: number }[] = [
    { keywords: ["hàn", "tiện", "cơ khí", "cửa sắt", "mái tôn", "nhôm kính", "inox"], ma: "25920", ten: "Gia công cơ khí; xử lý và tráng phủ kim loại", diem: 0.92 },
    { keywords: ["sửa chữa ô tô", "garage", "gara ô tô", "bảo dưỡng ô tô", "lốp ô tô"], ma: "45200", ten: "Bảo dưỡng, sửa chữa ô tô và xe có động cơ khác", diem: 0.95 },
    { keywords: ["sửa xe máy", "vá xe", "rửa xe máy", "thay dầu xe máy"], ma: "45403", ten: "Bảo dưỡng và sửa chữa mô tô, xe máy", diem: 0.90 },
    { keywords: ["may mặc", "may đo", "quần áo", "may gia công", "thời trang"], ma: "14100", ten: "May trang phục (trừ trang phục từ da lông thú)", diem: 0.88 },
    { keywords: ["quán cơm", "cơm tấm", "bún", "phở", "quán ăn", "nhà hàng"], ma: "56101", ten: "Quán ăn, nhà hàng, quán bún phở cơm phục vụ ăn uống tại chỗ", diem: 0.94 },
    { keywords: ["cà phê", "cafe", "trà sữa", "nước giải khát", "sinh tố"], ma: "56301", ten: "Quán cà phê, giải khát", diem: 0.91 },
    { keywords: ["tạp hóa", "bách hóa", "tiện ích", "bánh kẹo", "siêu thị mini"], ma: "47110", ten: "Bán lẻ lương thực, thực phẩm, đồ uống chiếm tỷ trọng lớn trong siêu thị mini", diem: 0.89 },
    { keywords: ["nhà thuốc", "quầy thuốc", "thuốc tây", "dược phẩm", "tân dược"], ma: "47721", ten: "Bán lẻ thuốc tân dược trong các cửa hàng chuyên doanh", diem: 0.96 },
    { keywords: ["điện thoại", "di động", "phụ kiện điện thoại", "smartphone"], ma: "47412", ten: "Bán lẻ điện thoại di động và thiết bị viễn thông", diem: 0.93 },
    { keywords: ["cắt tóc", "làm tóc", "uốn tóc", "gội đầu", "spa", "thẩm mỹ", "nail"], ma: "96310", ten: "Cắt tóc, làm đầu, gội đầu, làm móng và thẩm mỹ spa", diem: 0.95 },
    { keywords: ["vận tải", "chở hàng", "xe tải", "xe container", "giao hàng"], ma: "49339", ten: "Vận tải hàng hoá bằng xe tải khác", diem: 0.87 },
    { keywords: ["thức ăn chăn nuôi", "cám", "thức ăn gia súc", "thức ăn gia cầm"], ma: "46209", ten: "Bán buôn thức ăn và nguyên liệu làm thức ăn cho gia súc, gia cầm, thủy sản", diem: 0.91 }
  ];

  for (const rule of keywordRules) {
    if (rule.keywords.some(kw => d.includes(kw))) {
      return { ma: rule.ma, ten: rule.ten, diem: rule.diem };
    }
  }

  // Nếu không khớp từ khóa đặc biệt, tìm kiếm chuỗi trong danh mục chung
  for (const [code, name] of Object.entries(vsicRawData)) {
    if (code.length >= 4 && (name.toLowerCase().includes(d) || d.includes(name.toLowerCase()))) {
      return { ma: code, ten: name, diem: 0.75 };
    }
  }

  return null;
}

/**
 * Kiểm tra xem một hàng có phải là hàng tổng cộng / tổng số / bình quân hay không.
 */
export function isSummaryRow(row: any): boolean {
  if (!row || typeof row !== "object") return false;
  const summaryKeywords = ["tổng số", "tổng cộng", "cộng", "total", "bình quân", "toàn tỉnh", "toàn huyện", "toàn xã"];

  for (const val of Object.values(row)) {
    if (val === null || val === undefined) continue;
    const str = String(val).toLowerCase().trim();
    if (summaryKeywords.some(kw => str === kw || str.startsWith(kw + " ") || str.endsWith(" " + kw))) {
      return true;
    }
  }
  return false;
}

export function clearAllSectorsInVSIC(): void {}
export function clearAllParentsInVSIC(): void {}
