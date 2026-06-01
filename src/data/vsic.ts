export interface VSICSector {
  ma: string;
  ten: string;
  cap: number;
  cha: string | null;
}

export const vsicRawData: { [key: string]: string } = {
  "A": "NÔNG NGHIỆP, LÂM NGHIỆP VÀ THỦY SẢN",
  "B": "KHAI KHOÁNG",
  "C": "CÔNG NGHIỆP CHẾ BIẾN, CHẾ TẠO",
  "D": "SẢN XUẤT VÀ PHÂN PHỐI ĐIỆN, KHÍ ĐỐT, NƯỚC NÓNG, HƠI NƯỚC VÀ ĐIỀU HOÀ KHÔNG KHÍ",
  "E": "CUNG CẤP NƯỚC; HOẠT ĐỘNG QUẢN LÝ VÀ XỬ LÝ RÁC THẢI, NƯỚC THẢI",
  "F": "XÂY DỰNG",
  "G": "BÁN BUÔN VÀ BÁN LẺ; SỬA CHỮA Ô TÔ, MÔ TÔ, XE MÁY VÀ XE CÓ ĐỘNG CƠ KHÁC",
  "H": "VẬN TẢI KHO BÃI",
  "I": "DỊCH VỤ LƯU TRÚ VÀ ĂN UỐNG",
  "J": "THÔNG TIN VÀ TRUYỀN THÔNG",
  "K": "HOẠT ĐỘNG TÀI CHÍNH, NGÂN HÀNG VÀ BẢO HIỂM",
  "L": "HOẠT ĐỘNG KINH DOANH BẤT ĐỘNG SẢN",
  "M": "HOẠT ĐỘNG CHUYÊN MÔN, KHOA HỌC VÀ CÔNG NGHỆ",
  "N": "HOẠT ĐỘNG HÀNH CHÍNH VÀ DỊCH VỤ HỖ TRỢ",
  "O": "HOẠT ĐỘNG CỦA ĐẢNG CỘNG SẢN, TỔ CHỨC CHÍNH TRỊ - XÃ HỘI, QUẢN LÝ NHÀ NƯỚC, AN NINH QUỐC PHÒNG; BẢO ĐẢM XÃ HỘI BẮT BUỘC",
  "P": "GIÁO DỤC VÀ ĐÀO TẠO",
  "Q": "Y TẾ VÀ HOẠT ĐỘNG TRỢ GIÚP XÃ HỘI",
  "R": "NGHỆ THUẬT, VUI CHƠI VÀ GIẢI TRÍ",
  "S": "HOẠT ĐỘNG DỊCH VỤ KHÁC",
  "T": "HOẠT ĐỘNG LÀM THUÊ CÁC CÔNG VIỆC TRONG CÁC HỘ GIA ĐÌNH, SẢN XUẤT SẢN PHẨM VẬT CHẤT VÀ DỊCH VỤ TỰ TIÊU DÙNG CỦA HỘ GIA ĐÌNH",
  "U": "HOẠT ĐỘNG CỦA CÁC TỔ CHỨC VÀ CƠ QUAN QUỐC TẾ",
  "01": "NÔNG NGHIỆP VÀ HOẠT ĐỘNG DỊCH VỤ CÓ LIÊN QUAN",
  "011": "Trồng cây hàng năm",
  "0111": "Trồng lúa",
  "01110": "Trồng lúa",
  "0112": "Trồng ngô và cây lương thực có hạt khác",
  "01120": "Trồng ngô và cây lương thực có hạt khác",
  "0113": "Trồng cây lấy củ có chất bột",
  "01131": "Trồng khoai lang",
  "01132": "Trồng khoai tây",
  "01133": "Trồng sắn (khoai mỳ)",
  "01139": "Trồng cây lấy củ có chất bột khác",
  "0114": "Trồng cây mía",
  "01140": "Trồng cây mía",
  "0115": "Trồng cây thuốc lá, thuốc lào",
  "01150": "Trồng cây thuốc lá, thuốc lào",
  "0116": "Trồng cây lấy sợi",
  "01160": "Trồng cây lấy sợi",
  "0118": "Trồng rau, đậu các loại và trồng hoa",
  "01181": "Trồng rau các loại",
  "01182": "Trồng đậu các loại",
  "01183": "Trồng hoa các loại",
  "0119": "Trồng cây hàng năm khác",
  "01190": "Trồng cây hàng năm khác",
  "012": "Trồng cây lâu năm",
  "0121": "Trồng cây ăn quả",
  "01210": "Trồng cây ăn quả",
  "0122": "Trồng cây lấy quả chứa dầu",
  "01220": "Trồng cây lấy quả chứa dầu",
  "0123": "Trồng cây cà phê",
  "01230": "Trồng cây cà phê",
  "0124": "Trồng cây chè",
  "01240": "Trồng cây chè",
  "0125": "Trồng cây cao su",
  "01250": "Trồng cây cao su",
  "0126": "Trồng cây điều",
  "01260": "Trồng cây điều",
  "0127": "Trồng cây hồ tiêu",
  "01270": "Trồng cây hồ tiêu",
  "0128": "Trồng cây gia vị, cây dược liệu, cây hương liệu lâu năm",
  "0129": "Trồng cây lâu năm khác",
  "01290": "Trồng cây lâu năm khác",
  "013": "Nhân và chăm sóc giống cây trồng",
  "0130": "Nhân và chăm sóc giống cây trồng",
  "01300": "Nhân và chăm sóc giống cây trồng",
  "014": "Chăn nuôi",
  "0141": "Chăn nuôi trâu, bò",
  "01411": "Chăn nuôi trâu",
  "01412": "Chăn nuôi bò",
  "0142": "Chăn nuôi ngựa, lừa, la",
  "01420": "Chăn nuôi ngựa, lừa, la",
  "0143": "Chăn nuôi dê, cừu",
  "01431": "Chăn nuôi dê",
  "01432": "Chăn nuôi cừu",
  "0144": "Chăn nuôi lợn",
  "01440": "Chăn nuôi lợn",
  "0145": "Chăn nuôi gia cầm",
  "01450": "Chăn nuôi gia cầm",
  "0146": "Chăn nuôi khác",
  "01461": "Nuôi ong",
  "01462": "Nuôi tằm",
  "01469": "Chăn nuôi động vật khác chưa được cấu trúc",
  "0149": "Chăn nuôi hỗn hợp",
  "01490": "Chăn nuôi hỗn hợp",
  "015": "Hoạt động dịch vụ nông nghiệp",
  "016": "Săn bắt, bẫy và dịch vụ có liên quan",
  "0161": "Hoạt động dịch vụ trồng trọt",
  "0162": "Hoạt động dịch vụ chăn nuôi",
  "0163": "Hoạt động dịch vụ sau thu hoạch",
  "0164": "Xử lý hạt giống để nhân giống",
  "02": "LÂM NGHIỆP VÀ HOẠT ĐỘNG DỊCH VỤ CÓ LIÊN QUAN",
  "021": "Trồng rừng, chăm sóc rừng và ươm giống cây lâm nghiệp",
  "0210": "Trồng rừng, chăm sóc rừng và ươm giống cây lâm nghiệp",
  "02101": "Trồng rừng và chăm sóc rừng",
  "02102": "Ươm giống cây lâm nghiệp",
  "022": "Khai thác gỗ và lâm sản",
  "0221": "Khai thác gỗ",
  "0222": "Khai thác lâm sản ngoài gỗ",
  "023": "Thu nhặt lâm sản ngoài gỗ",
  "0230": "Thu nhặt lâm sản ngoài gỗ thực vật hoang dại",
  "024": "Hoạt động dịch vụ lâm nghiệp",
  "0240": "Hoạt động dịch vụ lâm nghiệp",
  "02400": "Hoạt động dịch vụ lâm nghiệp",
  "03": "KHAI THÁC, NUÔI TRỒNG THỦY SẢN",
  "031": "Khai thác thủy sản",
  "0311": "Khai thác thủy sản biển",
  "0312": "Khai thác thủy sản nội địa",
  "032": "Nuôi trồng thủy sản",
  "0321": "Nuôi trồng thủy sản biển",
  "0322": "Nuôi trồng thủy sản nội địa",
  "05": "KHAI THÁC THAN CỨNG VÀ THAN NON",
  "051": "Khai thác than cứng",
  "0511": "Khai thác than lộ thiên",
  "0512": "Khai thác than hầm lò",
  "052": "Khai thác than non",
  "0520": "Khai thác than non",
  "06": "KHAI THÁC DẦU THÔ VÀ KHÍ ĐỐT TỰ NHIÊN",
  "061": "Khai thác dầu thô",
  "0610": "Khai thác dầu thô",
  "062": "Khai thác khí dốt tự nhiên",
  "0620": "Khai thác khí đốt tự nhiên",
  "07": "KHAI THÁC QUẶNG KIM LOẠI",
  "071": "Khai thác quặng sắt",
  "0710": "Khai thác quặng sắt",
  "072": "Khai thác quặng kim loại màu",
  "08": "KHAI KHOÁNG KHÁC",
  "081": "Khai thác đá, cát, sỏi, đất sét",
  "089": "Khai khoáng khác chưa được phân vào đâu",
  "09": "HOẠT ĐỘNG DỊCH VỤ HỖ TRỢ KHAI KHOÁNG",
  "091": "Hoạt động dịch vụ hỗ trợ khai thác dầu thô và khí đốt tự nhiên",
  "099": "Hoạt động dịch vụ hỗ trợ khai khoáng khác",
  "10": "CHẾ BIẾN THỰC PHẨM",
  "101": "Chế biến, bảo quản thịt và các sản phẩm từ thịt",
  "102": "Chế biến, bảo quản thủy sản và các sản phẩm từ thủy sản",
  "103": "Chế biến và bảo quản rau quả",
  "104": "Sản xuất dầu, mỡ động, thực vật",
  "105": "Chế biến sữa và các sản phẩm từ sữa",
  "106": "Xay xát, sản xuất bột thô, tinh bột",
  "107": "Sản xuất các loại bánh kẹo, mứt",
  "108": "Sản xuất thức ăn gia súc, gia cầm và thủy sản",
  "109": "Sản xuất thực phẩm khác",
  "11": "SẢN XUẤT ĐỒ UỐNG",
  "110": "Sản xuất đồ uống",
  "12": "SẢN XUẤT SẢN PHẨM TỪ THUỐC LÁ",
  "120": "Sản xuất sản phẩm từ thuốc lá",
  "13": "DỆT",
  "131": "Sản xuất sợi, vải và hoàn thiện sản phẩm dệt",
  "139": "Sản xuất các loại hàng dệt khác",
  "14": "SẢN XUẤT TRANG PHỤC",
  "141": "Sản xuất trang phục (trừ trang phục từ da lông thú)",
  "142": "Sản xuất sản phẩm từ da lông thú",
  "143": "Sản xuất hàng may sẵn (trừ trang phục)",
  "15": "SẢN XUẤT DA, CÁC SẢN PHẨM LIÊN QUAN",
  "151": "Thuộc, sơ chế da; sản xuất vali, túi xách và các sản phẩm tương tự; sản xuất yên, cương",
  "152": "Sản xuất giày, dép",
  "16": "CHẾ BIẾN GỖ VÀ SẢN XUẤT SẢN PHẨM TỪ GỖ, TRE, NỨA (TRỪ GIƯỜNG, TỦ, BÀN, GHẾ); SẢN XUẤT CÁC SẢN PHẨM TỪ RƠM, RẠ VÀ VẬT LIỆU TẾT, BỆN",
  "161": "Cưa, xẻ, bào gỗ và bảo quản gỗ",
  "162": "Sản xuất sản phẩm từ gỗ, tre, nứa",
  "17": "SẢN XUẤT GIẤY VÀ SẢN PHẨM TỪ GIẤY",
  "170": "Sản xuất giấy và sản phẩm từ giấy",
  "18": "IN, SAO CHÉP BẢN GHI CÁC LOẠI",
  "181": "In, và các dịch vụ liên quan đến in",
  "182": "Sao chép bản ghi các loại",
  "19": "SẢN XUẤT THAN CỐC, SẢN PHẨM DẦU MỎ TINH CHẾ",
  "191": "Sản xuất than cốc",
  "192": "Sản xuất sản phẩm dầu mỏ tinh chế",
  "20": "SẢN XUẤT HÓA CHẤT VÀ SẢN PHẨM HÓA CHẤT",
  "201": "Sản xuất hoá chất cơ bản",
  "202": "Sản xuất sản phẩm hóa chất",
  "21": "SẢN XUẤT THUỐC, HOÁ DƯỢC VÀ DƯỢC LIỆU",
  "210": "Sản xuất thuốc, hoá dược và dược liệu",
  "22": "SẢN XUẤT SẢN PHẨM TỪ CAO SU VÀ PLASTIC",
  "221": "Sản xuất sản phẩm từ cao su",
  "222": "Sản xuất sản phẩm từ plastic",
  "23": "SẢN XUẤT SẢN PHẨM TỪ KHOÁNG PHI KIM LOẠI KHÁC",
  "231": "Sản xuất kính và sản phẩm từ kính",
  "239": "Sản xuất sản phẩm từ khoáng phi kim loại khác chưa được phân vào đâu",
  "24": "SẢN XUẤT KIM LOẠI",
  "241": "Sản xuất sắt, thép, gang",
  "242": "Sản xuất kim loại màu",
  "243": "Đúc kim loại",
  "25": "SẢN XUẤT SẢN PHẨM TỪ KIM LOẠI ĐÚC SẴN (TRỪ MÁY MÓC, THIẾT BỊ)",
  "251": "Sản xuất các cấu kiện kim loại, thùng, bể chứa và nồi hơi",
  "252": "Gia công cơ khí; xử lý và tráng phủ kim loại",
  "259": "Sản xuất các sản phẩm khác bằng kim loại",
  "26": "SẢN XUẤT SẢN PHẨM ĐIỆN TỬ, MÁY VI TÍNH VÀ SẢN PHẨM QUANG HỌC",
  "261": "Sản xuất linh kiện điện tử",
  "262": "Sản xuất máy vi tính và thiết bị ngoại vi của máy vi tính",
  "263": "Sản xuất thiết bị truyền thông",
  "264": "Sản xuất sản phẩm điện tử dân dụng",
  "265": "Sản xuất thiết bị đo lường, kiểm tra, định hướng và điều khiển",
  "266": "Sản xuất đồng hồ",
  "267": "Sản xuất thiết bị và dụng cụ quang học",
  "268": "Sản xuất băng, đĩa từ tính và quang học",
  "27": "SẢN XUẤT THIỆT BỊ CHIẾU SÁNG VÀ THIẾT BỊ ĐIỆN KHÁC",
  "271": "Sản xuất mô tơ, máy phát, biến thế điện, thiết bị phân phối và điều khiển điện",
  "272": "Sản xuất pin và ắc quy",
  "273": "Sản xuất dây và thiết bị dây dẫn điện các loại",
  "274": "Sản xuất thiết bị chiếu sáng",
  "275": "Sản xuất đồ điện dân dụng",
  "279": "Sản xuất thiết bị điện khác",
  "28": "SẢN XUẤT MÁY MÓC, THIẾT BỊ CHƯA ĐƯỢC PHÂN VÀO ĐÂU",
  "281": "Sản xuất máy móc thông dụng",
  "282": "Sản xuất máy móc chuyên dụng",
  "29": "SẢN XUẤT Ô TÔ VÀ XE CÓ ĐỘNG CƠ KHÁC",
  "291": "Sản xuất ô tô và động cơ ô tô",
  "292": "Sản xuất thân xe ô tô và rơ moóc",
  "293": "Sản xuất phụ tùng và bộ phận phụ trợ cho xe ô tô và xe có động cơ khác",
  "30": "SẢN XUẤT PHƯƠNG TIỆN VẬN TẢI KHÁC",
  "301": "Đóng tàu và thuyền",
  "302": "Sản xuất đầu máy xe lửa, xe điện và toa xe",
  "303": "Sản xuất máy bay, tàu vũ trụ và máy móc liên quan",
  "309": "Sản xuất phương tiện vận tải khác chưa được phân vào đâu",
  "31": "SẢN XUẤT GIƯỜNG, TỦ, BÀN, GHẾ",
  "310": "Sản xuất giường, tủ, bàn, ghế",
  "3100": "Sản xuất giường, tủ, bàn, ghế",
  "31001": "Sản xuất giường, tủ, bàn, ghế bằng gỗ",
  "31002": "Sản xuất giường, tủ, bàn, ghế bằng kim loại",
  "31009": "Sản xuất giường, tủ, bàn, ghế bằng vật liệu khác",
  "32": "CÔNG NGHIỆP CHẾ BIẾN, CHẾ TẠO KHÁC",
  "321": "Sản xuất đồ kim hoàn, đồ giả kim hoàn và các chi tiết liên quan",
  "322": "Sản xuất nhạc cụ",
  "323": "Sản xuất dụng cụ thể dục, thể thao",
  "324": "Sản xuất đồ chơi, trò chơi",
  "325": "Sản xuất thiết bị, dụng cụ y tế, nha khoa",
  "329": "Sản xuất khác chưa được phân vào đâu",
  "33": "SỬA CHỮA, BẢO DƯỠNG VÀ LẮP ĐẶT MÁY MÓC VÀ THIẾT BỊ",
  "331": "Sửa chữa, bảo dưỡng máy móc, thiết bị",
  "332": "Lắp đặt máy móc và thiết bị công nghiệp",
  "35": "SẢN XUẤT VÀ PHÂN PHỐI ĐIỆN, KHÍ ĐỐT, NƯỚC NÓNG, HƠI NƯỚC VÀ ĐIỀU HOÀ KHÔNG KHÍ",
  "351": "Sản xuất, truyền tải và phân phối điện",
  "3511": "Sản xuất điện",
  "35111": "Thủy điện",
  "35112": "Nhiệt điện",
  "35113": "Điện gió",
  "35114": "Điện mặt trời",
  "3512": "Truyền tải và phân phối điện",
  "352": "Sản xuất khí đốt; phân phối nhiên liệu khí bằng đường ống",
  "353": "Sản xuất, phân phối hơi nước, nước nóng, điều hoà không khí và sản xuất nước đá",
  "36": "KHAI THÁC, XỬ LÝ VÀ CUNG CẤP NƯỚC",
  "360": "Khai thác, xử lý và cung cấp nước",
  "37": "THOÁT NƯỚC VÀ XỬ LÝ NƯỚC THẢI",
  "370": "Thoát nước và xử lý nước thải",
  "38": "HOẠT ĐỘNG THU GOM, XỬ LÝ VÀ TIÊU HUỶ RÁC THẢI; TÁI CHẾ PHẾ LIỆU",
  "381": "Thu gom rác thải",
  "382": "Xử lý và tiêu huỷ rác thải",
  "383": "Tái chế phế liệu",
  "39": "XỬ LÝ Ô NHIỄM VÀ HOẠT ĐỘNG QUẢN LÝ CHẤT THẢI KHÁC",
  "390": "Xử lý ô nhiễm và hoạt động quản lý chất thải khác",
  "41": "XÂY DỰNG NHÀ CÁC LOẠI",
  "410": "Xây dựng nhà các loại",
  "4100": "Xây dựng nhà các loại",
  "41000": "Xây dựng nhà các loại",
  "42": "XÂY DỰNG CÔNG TRÌNH KỸ THUẬT DÂN DỤNG",
  "421": "Xây dựng công trình đường sắt và đường bộ",
  "4210": "Xây dựng công trình đường sắt và đường bộ",
  "422": "Xây dựng công trình công ích",
  "429": "Xây dựng công trình kỹ thuật dân dụng khác",
  "43": "HOẠT ĐỘNG XÂY DỰNG CHUYÊN DỤNG",
  "431": "Phá dỡ và chuẩn bị mặt bằng",
  "432": "Lắp đặt hệ thống xây dựng",
  "433": "Hoàn thiện công trình xây dựng",
  "439": "Hoạt động xây dựng chuyên dụng khác",
  "45": "BÁN, SỬA CHỮA Ô TÔ, MÔ TÔ, XE MÁY VÀ XE CÓ ĐỘNG CƠ KHÁC",
  "451": "Bán ô tô và xe có động cơ khác",
  "452": "Bảo dưỡng, sửa chữa ô tô và xe có động cơ khác",
  "453": "Bán phụ tùng và các bộ phận phụ trợ của ô tô và xe có động cơ khác",
  "454": "Bán, bảo dưỡng, sửa chữa mô tô, xe máy, phụ tùng và các bộ phận phụ trợ",
  "46": "BÁN BUÔN (TRỪ Ô TÔ, MÔ TÔ, XE MÁY VÀ XE CÓ ĐỘNG CƠ KHÁC)",
  "461": "Đại lý, môi giới, đấu giá",
  "462": "Bán buôn nông, lâm sản nguyên liệu và động vật sống",
  "463": "Bán buôn thực phẩm, đồ uống và sản phẩm thuốc lá, thuốc lào",
  "464": "Bán buôn đồ dùng gia đình",
  "465": "Bán buôn máy móc, thiết bị và phụ tùng máy",
  "466": "Bán buôn chuyên doanh khác",
  "47": "BÁN LẺ (TRỪ Ô TÔ, MÔ TÔ, XE MÁY VÀ XE CÓ ĐỘNG CƠ KHÁC)",
  "471": "Bán lẻ trong các cửa hàng kinh doanh tổng hợp",
  "472": "Bán lẻ lương thực, thực phẩm, đồ uống, thuốc lá chiếm tỷ trọng lớn",
  "473": "Bán lẻ nhiên liệu động cơ trong các cửa hàng chuyên doanh",
  "474": "Bán lẻ thiết bị công nghệ thông tin trong các cửa hàng chuyên doanh",
  "475": "Bán lẻ đồ dùng gia đình khác trong các cửa hàng chuyên doanh",
  "476": "Bán lẻ hàng văn hoá, giải trí trong các cửa hàng chuyên doanh",
  "477": "Bán lẻ hàng hoá khác trong các cửa hàng chuyên doanh",
  "478": "Bán lẻ lưu động hoặc bán lẻ theo yêu cầu đặt hàng",
  "479": "Bán lẻ hình thức khác (Internet, bưu điện, qua tivi)",
  "49": "VẬN TẢI ĐƯỜNG SẮT, ĐƯỜNG BỘ VÀ VẬN TẢI ĐƯỜNG ỐNG",
  "491": "Vận tải đường sắt",
  "492": "Vận tải hành khách đường bộ trong nội, ngoại thành",
  "493": "Vận tải hàng hóa đường bộ",
  "494": "Vận tải đường ống",
  "50": "VẬN TẢI ĐƯỜNG THỦY",
  "501": "Vận tải ven biển và viễn dương",
  "502": "Vận tải đường thủy nội địa",
  "51": "VẬN TẢI HÀNG KHÔNG",
  "511": "Vận tải hành khách hàng không",
  "512": "Vận tải hàng hóa hàng không",
  "52": "KHO BÃI VÀ CÁC HOẠT ĐỘNG HỖ TRỢ CHO VẬN TẢI",
  "521": "Kho bãi và lưu giữ hàng hoá",
  "522": "Hoạt động dịch vụ hỗ trợ trực tiếp cho vận tải",
  "53": "BƯU CHÍNH VÀ CHUYỂN PHÁT",
  "531": "Bưu chính",
  "532": "Chuyển phát",
  "55": "DỊCH VỤ LƯU TRÚ",
  "551": "Dịch vụ lưu trú ngắn ngày",
  "559": "Cơ sở lưu trú khác",
  "56": "DỊCH VỤ ĂN UỐNG",
  "561": "Nhà hàng và các dịch vụ ăn uống phục vụ lưu động",
  "562": "Dịch vụ phục vụ đồ uống",
  "563": "Cung cấp dịch vụ ăn uống theo hợp đồng không thường xuyên",
  "58": "HOẠT ĐỘNG XUẤT BẢN",
  "581": "Hoạt động xuất bản",
  "582": "Xuất bản phần mềm",
  "59": "HOẠT ĐỘNG ĐIỆN ẢNH, SẢN XUẤT CHƯƠNG TRÌNH TRUYỀN HÌNH, GHI ÂM VÀ XUẤT BẢN ÂM NHẠC",
  "591": "Hoạt động điện ảnh, sản xuất chương trình truyền hình",
  "592": "Hoạt động ghi âm và xuất bản âm nhạc",
  "60": "HOẠT ĐỘNG PHÁT THANH, TRUYỀN HÌNH",
  "601": "Hoạt động phát thanh",
  "602": "Hoạt động truyền hình",
  "61": "VIỄN THÔNG",
  "611": "Hoạt động viễn thông có dây",
  "612": "Hoạt động viễn thông không dây",
  "613": "Hoạt động viễn thông vệ tinh",
  "619": "Hoạt động viễn thông khác",
  "62": "LẬP TRÌNH MÁY VI TÍNH, DỊCH VỤ TƯ VẤN VÀ CÁC HOẠT ĐỘNG KHÁC LIÊN QUAN ĐẾN MÁY VI TÍNH",
  "620": "Lập trình máy vi tính, tư vấn và hoạt động liên quan",
  "63": "HOẠT ĐỘNG DỊCH VỤ THÔNG TIN",
  "631": "Xử lý dữ liệu, cho thuê và các hoạt động liên quan",
  "639": "Hoạt động dịch vụ thông tin khác",
  "64": "HOẠT ĐỘNG DỊCH VỤ TÀI CHÍNH (TRỪ BẢO HIỂM VÀ BẢO HIỂM XÃ HỘI)",
  "641": "Trung gian tài chính",
  "642": "Hoạt động của các công ty nắm giữ tài sản",
  "643": "Hoạt động của các quỹ tín thác, các quỹ và tổ chức tài chính tương tự",
  "649": "Hoạt động dịch vụ tài chính khác chưa được phân vào đâu",
  "65": "BẢO HIỂM, TÁI BẢO HIỂM VÀ BẢO HIỂM XÃ HỘI (TRỪ BẢO ĐẢM XÃ HỘI)",
  "651": "Bảo hiểm",
  "652": "Tái bảo hiểm",
  "653": "Bảo hiểm xã hội",
  "66": "HOẠT ĐỘNG HỖ TRỢ DỊCH VỤ TÀI CHÍNH, BẢO HIỂM VÀ BẢO HIỂM XÃ HỘI",
  "661": "Hoạt động hỗ trợ dịch vụ tài chính",
  "662": "Hoạt động hỗ trợ bảo hiểm và bảo hiểm xã hội",
  "663": "Hoạt động quản lý quỹ",
  "68": "HOẠT ĐỘNG KINH DOANH BẤT ĐỘNG SẢN",
  "681": "Kinh doanh bất động sản, quyền sử dụng đất thuộc chủ sở hữu hoặc đi thuê",
  "682": "Tư vấn, môi giới, đấu giá bất động sản, đấu giá quyền sử dụng đất",
  "69": "HOẠT ĐỘNG PHÁP LUẬT, KẾ TOÁN VÀ KIỂM TOÁN",
  "691": "Hoạt động pháp luật",
  "692": "Hoạt động kế toán, kiểm toán và tư vấn thuế",
  "70": "HOẠT ĐỘNG CỦA TRỤ SỞ VĂN PHÒNG; HOẠT ĐỘNG TƯ VẤN QUẢN LÝ",
  "701": "Hoạt động của trụ sở văn phòng",
  "702": "Hoạt động tư vấn quản lý",
  "71": "HOẠT ĐỘNG KIẾN TRÚC, KIỂM TRA VÀ PHÂN TÍCH KỸ THUẬT",
  "711": "Hoạt động kiến trúc và tư vấn kỹ thuật có liên quan",
  "712": "Kiểm tra và phân tích kỹ thuật",
  "72": "NGHIÊN CỨU KHOA HỌC VÀ PHÁT TRIỂN",
  "721": "Nghiên cứu khoa học và phát triển trong khoa học tự nhiên và kỹ thuật",
  "722": "Nghiên cứu khoa học và phát triển trong khoa học xã hội và nhân văn",
  "73": "QUẢNG CÁO VÀ NGHIÊN CỨU THỊ TRƯỜNG",
  "731": "Quảng cáo",
  "732": "Nghiên cứu thị trường và thăm dò dư luận",
  "74": "HOẠT ĐỘNG CHUYÊN MÔN, KHOA HỌC VÀ CÔNG NGHỆ KHÁC",
  "741": "Thiết kế chuyên dụng",
  "742": "Hoạt động nhiếp ảnh",
  "749": "Hoạt động chuyên môn, khoa học và công nghệ khác chưa được phân vào đâu",
  "75": "HOẠT ĐỘNG THÚ Y",
  "750": "Hoạt động thú y",
  "77": "CHO THUÊ MÁY MÓC, THIẾT BỊ; CHO THUÊ ĐỒ DÙNG CÁ NHÂN VÀ GIA ĐÌNH; CHO THUÊ TÀI SẢN VÔ HÌNH",
  "771": "Cho thuê xe có động cơ",
  "772": "Cho thuê đồ dùng cá nhân và gia đình",
  "773": "Cho thuê máy móc, thiết bị và đồ dùng hữu hình",
  "774": "Cho thuê tài sản vô hình không có bản quyền",
  "78": "HOẠT ĐỘNG DỊCH VỤ LAO ĐỘNG VÀ VIỆC LÀM",
  "781": "Hoạt động của các trung tâm tư vấn, giới thiệu và môi giới lao động",
  "782": "Cung ứng lao động tạm thời",
  "783": "Cung ứng và quản lý nguồn lao động",
  "79": "HOẠT ĐỘNG CỦA CÁC ĐẠI LÝ DU LỊCH, KINH DOANH TOUR DU LỊCH VÀ CÁC DỊCH VỤ HỖ TRỢ",
  "791": "Hoạt động của các đại lý du lịch, kinh doanh tour du lịch",
  "799": "Dịch vụ đặt chỗ và dịch vụ hỗ trợ liên quan",
  "80": "HOẠT ĐỘNG ĐIỀU TRA BẢO ĐẢM AN TOÀN",
  "801": "Hoạt động dịch vụ bảo vệ cá nhân",
  "802": "Hoạt động dịch vụ hệ thống an toàn",
  "803": "Hoạt động điều tra",
  "81": "HOẠT ĐỘNG DỊCH VỤ VỆ SINH NHÀ CỬA, CÔNG TRÌNH VÀ CẢNH QUAN",
  "811": "Vệ sinh chung nhà cửa",
  "812": "Vệ sinh công nghiệp và các công trình chuyên biệt",
  "813": "Dịch vụ chăm sóc và duy trì cảnh quan",
  "82": "HOẠT ĐỘNG HÀNH CHÍNH, HỖ TRỢ VĂN PHÒNG VÀ CÁC HOẠT ĐỘNG HỖ TRỢ KINH DOANH KHÁC",
  "821": "Hoạt động hành chính và hỗ trợ văn phòng",
  "822": "Hoạt động dịch vụ đóng gói",
  "823": "Hoạt động tổ chức giới thiệu và xúc tiến thương mại",
  "829": "Hoạt động dịch vụ hỗ trợ kinh doanh khác chưa được phân vào đâu",
  "84": "HOẠT ĐỘNG CỦA ĐẢNG CỘNG SẢN, TỔ CHỨC CHÍNH TRỊ - XÃ HỘI, QUẢN LÝ NHÀ NƯỚC, AN NINH QUỐC PHÒNG; BẢO ĐẢM XÃ HỘI BẮT BUỘC",
  "841": "Quản lý nhà nước và chính sách kinh tế - xã hội",
  "842": "Dịch vụ cho toàn xã hội",
  "843": "Bảo đảm xã hội bắt buộc",
  "85": "GIÁO DỤC VÀ ĐÀO TẠO",
  "851": "Giáo dục mầm non",
  "852": "Giáo dục phổ thông",
  "853": "Giáo dục nghề nghiệp",
  "854": "Đào tạo đại học và sau đại học",
  "855": "Các hình thức giáo dục khác",
  "856": "Dịch vụ hỗ trợ giáo dục",
  "86": "HOẠT ĐỘNG Y TẾ",
  "861": "Hoạt động của các bệnh viện, trạm y tế",
  "862": "Hoạt động của các phòng khám đa khoa, chuyên khoa và nha khoa",
  "869": "Hoạt động y tế khác",
  "87": "HOẠT ĐỘNG CHĂM SÓC, NUÔI DƯỠNG TẬP TRUNG",
  "871": "Hoạt động chăm sóc tập trung cho người có công, cao tuổi, khuyết tật",
  "872": "Hoạt động chăm sóc tập trung cho người nghiện ma túy, HIV/AIDS",
  "873": "Hoạt động chăm sóc tập trung cho người tâm thần",
  "879": "Hoạt động chăm sóc, nuôi dưỡng tập trung khác",
  "88": "HOẠT ĐỘNG TRỢ GIÚP XÃ HỘI KHÔNG TẬP TRUNG",
  "881": "Hoạt động trợ giúp xã hội không tập trung cho người cao tuổi, khuyết tật",
  "889": "Hoạt động trợ giúp xã hội không tập trung khác",
  "90": "HOẠT ĐỘNG SÁNG TÁC, NGHỆ THUẬT VÀ GIẢI TRÍ",
  "900": "Hoạt động sáng tác, nghệ thuật và giải trí",
  "91": "HOẠT ĐỘNG CỦA THƯ VIỆN, LƯU TRỮ, BẢO TÀNG VÀ CÁC HOẠT ĐỘNG VĂN HÓA KHÁC",
  "910": "Hoạt động của thư viện, lưu trữ, bảo tàng và hoạt động văn hóa khác",
  "92": "HOẠT ĐỘNG XỔ SỐ, CÁ CƯỢC VÀ ĐÁNH BẠC",
  "920": "Hoạt động xổ số, cá cược và đánh bạc",
  "93": "HOẠT ĐỘNG THỂ THAO, VUI CHƠI VÀ GIẢI TRÍ",
  "931": "Hoạt động thể thao",
  "932": "Hoạt động vui chơi giải trí khác",
  "94": "HOẠT ĐỘNG CỦA CÁC TỔ CHỨC HIỆP HỘI",
  "941": "Hoạt động của các tổ chức kinh doanh, chủ sử dụng và tổ chức chuyên môn",
  "942": "Hoạt động của các công đoàn",
  "949": "Hoạt động của các tổ chức hiệp hội khác chưa được phân vào đâu",
  "95": "SỬA CHỮA MÁY VI TÍNH, ĐỒ DÙNG CÁ NHÂN VÀ GIA ĐÌNH",
  "951": "Sửa chữa máy vi tính và thiết bị liên lạc",
  "952": "Sửa chữa đồ dùng cá nhân và gia đình",
  "96": "HOẠT ĐỘNG DỊCH VỤ PHỤC VỤ CÁ NHÂN KHÁC",
  "961": "Giặt là, làm sạch các sản phẩm dệt và da",
  "962": "Cắt tóc, làm đầu, gội đầu",
  "963": "Dịch vụ tang lễ và các hoạt động liên quan",
  "969": "Hoạt động dịch vụ phục vụ cá nhân khác còn lại",
  "97": "HOẠT ĐỘNG LÀM THUÊ CÔNG VIỆC GIA ĐÌNH TRONG CÁC HỘ GIA ĐÌNH",
  "970": "Hoạt động làm thuê công việc gia đình trong các hộ gia đình",
  "98": "HOẠT ĐỘNG SẢN XUẤT SẢN PHẨM VẬT CHẤT VÀ DỊCH VỤ TỰ TIÊU DÙNG CỦA HỘ GIA ĐÌNH",
  "981": "Hoạt động sản xuất sản phẩm vật chất tự tiêu dùng của hộ gia đình",
  "982": "Hoạt động sản xuất sản phẩm dịch vụ tự tiêu dùng của hộ gia đình",
  "99": "HOẠT ĐỘNG CỦA CÁC TỔ CHỨC VÀ CƠ QUAN QUỐC TẾ",
  "990": "Hoạt động của các tổ chức và cơ quan quốc tế"
};

// Chuẩn hóa mã ngành, loại bỏ .0 của Excel và padding số 0 ở đầu
export function normalizeSectorCode(code: any): string {
  if (code === null || code === undefined) return "";
  let str = String(code).trim();
  if (str.endsWith(".0")) {
    str = str.slice(0, -2);
  }
  str = str.replace(/[^a-zA-Z0-9]/g, "");
  if (!str) return "";

  // Nếu là chuỗi số, tự động padding số 0 ở đầu nếu độ dài thiếu so với chuẩn cấp
  if (/^\d+$/.test(str)) {
    // Nếu độ dài là 1: padding thành cấp 2 (ví dụ: '1' -> '01')
    if (str.length === 1) {
      str = "0" + str;
    } 
    // Nếu độ dài là 4: thường là cấp 5 dính mất số 0 ở đầu (ví dụ: '1110' -> '01110')
    else if (str.length === 4) {
      // Chúng ta kiểm tra xem nếu padding 0 ở đầu có tạo thành mã cấp 5 hợp lệ không
      const possibleCap5 = "0" + str;
      if (vsicRawData[possibleCap5]) {
        str = possibleCap5;
      }
    }
  }
  return str;
}

// Khắc phục lỗi đứt gãy liên kết cha của mã ngành cấp 2 và cấp 1
export function getParentSectorCode(code: string): string | null {
  const cleanCode = code.trim();
  if (cleanCode.length <= 1) return null; // Cấp 1 (A-U) không có cha
  
  if (cleanCode.length === 2) {
    const num = parseInt(cleanCode, 10);
    if (isNaN(num)) return null;
    
    // Bản đồ phân nhóm cấp 2 sang ngành cấp 1 (VSIC Việt Nam)
    if (num >= 1 && num <= 3) return "A";
    if (num >= 5 && num <= 9) return "B";
    if (num >= 10 && num <= 33) return "C";
    if (num === 35) return "D";
    if (num >= 36 && num <= 39) return "E";
    if (num >= 41 && num <= 43) return "F";
    if (num >= 45 && num <= 47) return "G";
    if (num >= 49 && num <= 53) return "H";
    if (num >= 55 && num <= 56) return "I";
    if (num >= 58 && num <= 63) return "J";
    if (num >= 64 && num <= 66) return "K";
    if (num === 68) return "L";
    if (num >= 69 && num <= 75) return "M";
    if (num >= 77 && num <= 82) return "N";
    if (num === 84) return "O";
    if (num === 85) return "P";
    if (num >= 86 && num <= 88) return "Q";
    if (num >= 90 && num <= 93) return "R";
    if (num >= 94 && num <= 96) return "S";
    if (num >= 97 && num <= 98) return "T";
    if (num === 99) return "U";
    return null;
  }
  
  // Cấp 5 (5 chữ số) -> cấp 4 (4 chữ số) -> cấp 3 (3 chữ số) -> cấp 2 (2 chữ số)
  return cleanCode.slice(0, cleanCode.length - 1);
}

// Lấy thông tin cấp của ngành nghề dựa trên cấu trúc mã
export function getSectorLevel(code: string): number {
  const cleanCode = code.trim();
  if (!cleanCode) return 0;
  if (/^[a-zA-Z]$/.test(cleanCode)) return 1;
  return cleanCode.length; // 2, 3, 4, 5 tương ứng cấp 2, 3, 4, 5
}

// Lấy đầy đủ thông tin phân cấp (Ancestors) của một ngành nghề lên cấp 1
export interface ParentHierarchy {
  ma: string;
  ten: string;
  cap: number;
}

// Quy chuẩn tên ngành cấp 1 và cấp 2 theo yêu cầu hiển thị
export function formatVSICName(code: string, rawName: string): string {
  const cleanCode = code.trim();
  if (cleanCode === "C" || cleanCode === "c") return "Công nghiệp chế biến, chế tạo";
  if (cleanCode === "10") return "Sản xuất, chế biến thực phẩm";
  if (cleanCode === "A" || cleanCode === "a") return "Nông nghiệp, lâm nghiệp và thủy sản";
  
  // Chuẩn hóa chữ hoa đầu tiên cho nhãn hiển thị nếu tên gốc viết hoa hoàn toàn
  if (rawName === rawName.toUpperCase() && rawName.length > 3) {
    return rawName.charAt(0) + rawName.slice(1).toLowerCase();
  }
  return rawName;
}

export function getSectorHierarchy(code: string): { [key: string]: ParentHierarchy | null } {
  const result: { [key: string]: ParentHierarchy | null } = {
    "1": null,
    "2": null,
    "3": null,
    "4": null,
    "5": null
  };

  let currentCode = normalizeSectorCode(code);
  while (currentCode) {
    const level = getSectorLevel(currentCode);
    let ten = vsicRawData[currentCode];
    if (ten) {
      ten = formatVSICName(currentCode, ten);
      result[String(level)] = {
        ma: currentCode,
        ten: ten,
        cap: level
      };
    } else {
      // Trong trường hợp mã ngành chi tiết (ví dụ: cấp 5) không tồn tại trong danh mục,
      // ta tìm kiếm ngược lên các tổ tiên phân cấp gần nhất có tên được định nghĩa
      let fallbackTen = `[Ngành cấp ${level} chưa định nghĩa]`;
      let checkCode = currentCode;
      while (checkCode.length > 1) {
        checkCode = checkCode.slice(0, checkCode.length - 1);
        const pTen = vsicRawData[checkCode];
        if (pTen) {
          fallbackTen = formatVSICName(checkCode, pTen);
          break;
        }
      }
      result[String(level)] = {
        ma: currentCode,
        ten: fallbackTen,
        cap: level
      };
    }
    
    // Tìm mã cha của level hiện tại
    const chaCode = getParentSectorCode(currentCode);
    if (!chaCode || chaCode === currentCode) break;
    currentCode = chaCode;
  }

  return result;
}

// Tìm kiếm mã ngành thủ công / thông minh
export function searchVSICSectors(query: string, maxResults = 10): VSICSector[] {
  const normalizedQuery = query.toLowerCase().trim();
  if (!normalizedQuery) return [];

  const results: VSICSector[] = [];
  const keys = Object.keys(vsicRawData);

  for (const ma of keys) {
    const ten = vsicRawData[ma];
    const matchMa = ma.toLowerCase().includes(normalizedQuery);
    const matchTen = ten.toLowerCase().includes(normalizedQuery);

    if (matchMa || matchTen) {
      results.push({
        ma,
        ten,
        cap: getSectorLevel(ma),
        cha: getParentSectorCode(ma)
      });
      if (results.length >= maxResults * 2) break; // Lấy dôi ra để sắp xếp
    }
  }

  // Sắp xếp: ưu tiên trùng khớp mã chính xác hoặc xuất hiện ở đầu
  return results
    .sort((a, b) => {
      const aMaIdx = a.ma.toLowerCase().indexOf(normalizedQuery);
      const bMaIdx = b.ma.toLowerCase().indexOf(normalizedQuery);
      if (aMaIdx === 0 && bMaIdx !== 0) return -1;
      if (bMaIdx === 0 && aMaIdx !== 0) return 1;
      
      const aTenIdx = a.ten.toLowerCase().indexOf(normalizedQuery);
      const bTenIdx = b.ten.toLowerCase().indexOf(normalizedQuery);
      if (aTenIdx === 0 && bTenIdx !== 0) return -1;
      if (bTenIdx === 0 && aTenIdx !== 0) return 1;

      return a.ma.length - b.ma.length; // Ưu tiên mã ngắn hơn (cấp cao hơn)
    })
    .slice(0, maxResults);
}

// Từ khóa phân biệt nhanh các chương ngành chính (Keywords) để giúp khớp ngành không cần AI
export const kwMapping: { [key: string]: string[] } = {
  "A": ["nông nghiệp", "trồng trọt", "chăn nuôi", "lúa", "ngô", "sắn", "khoai", "rau", "quả", "cà phê", "chè", "cao su", "điều", "tiêu", "gia cầm", "lợn", "trâu", "bò", "rừng", "lâm sản", "nuôi trồng thủy sản", "khai thác thủy sản", "đánh bắt thủy sản", "cá", "tôm", "hải sản"],
  "B": ["khai thác", "than", "quặng", "quặng sắt", "dầu thô", "khí đốt", "đá", "cát", "sỏi", "đất sét", "khai khoáng"],
  "C": ["sản xuất", "chế biến", "bánh kẹo", "nước ngọt", "bia", "rượu", "thuốc lá", "dệt", "vải", "may mặc", "quần áo", "giày", "dép", "da", "gỗ", "tre", "giấy", "in ấn", "hóa chất", "dược liệu", "thuốc", "cao su", "nhựa", "plastic", "kính", "xi măng", "gốm", "sắt", "thép", "gang", "đúc kim loại", "gia công cơ khí", "linh kiện", "điện tử", "máy tính", "điện thoại", "thiết bị chiếu sáng", "mô tơ", "ắc quy", "ô tô", "tàu", "máy bay", "bàn ghế", "giường tủ", "vàng bạc", "nhạc cụ", "đồ chơi"],
  "D": ["điện", "phát điện", "nhiệt điện", "thủy điện", "điện gió", "điện mặt trời", "khí đốt", "hơi nước", "điều hoà", "nước nóng"],
  "E": ["cung cấp nước", "nước sạch", "xử lý rác", "nước thải", "thu gom rác", "tái chế", "phế liệu", "môi trường"],
  "F": ["xây dựng", "xây nhà", "thi công", "đường bộ", "cầu đường", "phá dỡ", "chuẩn bị mặt bằng", "hoàn thiện công trình", "lắp đặt điện nước"],
  "G": ["bán buôn", "bán lẻ", "cho thuê ô tô", "mô tô", "xe máy", "bán lẻ thiết bị", "cửa hàng tạp hóa", "siêu thị", "chợ", "thương mại"],
  "H": ["vận tải", "shiper", "gửi hàng", "taxi", "xe buýt", "xe khách", "xe tải", "đường sắt", "kho bãi", "bưu chính", "chuyển phát", "giao hàng"],
  "I": ["nhà hàng", "khách sạn", "nhà nghỉ", "homestay", "quán ăn", "quán cà phê", "quán nước", "ẩm thực", "lưu trú", "ăn uống"],
  "J": ["xuất bản", "phần mềm", "điện ảnh", "truyền hình", "phát thanh", "viễn thông", "internet", "lập trình", "máy vi tính", "website", "công nghệ thông tin", "cntt", "portal"],
  "K": ["tài chính", "ngân hàng", "bảo hiểm", "tín dụng", "quỹ đầu tư", "chứng khoán"],
  "L": ["bất động sản", "nhà đất", "cho thuê nhà", "môi giới nhà đất", "địa ốc"],
  "M": ["pháp luật", "luật sư", "kế toán", "kiểm toán", "tư vấn thuế", "kiến trúc", "thiết kế", "kiểm tra kỹ thuật", "nghiên cứu khoa học", "quảng cáo", "marketing", "nhiếp ảnh", "thú y"],
  "N": ["cho thuê máy móc", "du lịch", "tour du lịch", "bảo vệ", "vệ sinh", "dọn dẹp", "hành chính", "văn phòng", "đóng gói", "giới thiệu việc làm"],
  "O": ["quản lý nhà nước", "an ninh", "quốc phòng", "bảo đảm xã hội"],
  "P": ["giáo dục", "đào tạo", "trường học", "mầm non", "tiểu học", "thpt", "thcs", "đại học", "dạy nghề", "trung tâm tiếng anh"],
  "Q": ["y tế", "bệnh viện", "phòng khám", "nha khoa", "trợ giúp xã hội", "viện dưỡng lão"],
  "R": ["sáng tác", "nghệ thuật", "giải trí", "thư viện", "bảo tàng", "xổ số", "cá cược", "đánh bạc", "thể thao", "sân bóng", "gym", "khu vui chơi"],
  "S": ["hiệp hội", "công đoàn", "sửa chữa máy tính", "sửa điện thoại", "giặt là", "làm đầu", "cắt tóc", "gội đầu", "spa", "tang lễ", "massage", "tắm hơi"],
  "T": ["làm thuê", "giúp việc", "hộ gia đình"],
  "U": ["tổ chức quốc tế", "cơ quan quốc tế", "un", "unesco", "embassy", "đại sứ quán"]
};

// Thuật toán so khớp thông minh bằng từ khóa kết hợp khoảng cách từ (TF-IDF / Trùng khớp một phần)
export function smartSuggestSectorByDescription(description: string): { ma: string; ten: string; diem: number } | null {
  const text = description.toLowerCase().trim();
  if (!text) return null;

  // 1. Phân luồng theo từ khóa ngành Cấp 1 để thu hẹp phạm vi
  let bestMainSector: string | null = null;
  let maxKeywordCount = 0;

  for (const mainSector of Object.keys(kwMapping)) {
    const kws = kwMapping[mainSector];
    let count = 0;
    for (const kw of kws) {
      if (text.includes(kw)) {
        // Cộng điểm ưu tiên cho từ khóa có độ dài lớn hơn
        count += kw.split(" ").length;
      }
    }
    if (count > maxKeywordCount) {
      maxKeywordCount = count;
      bestMainSector = mainSector;
    }
  }

  // 2. Tra cứu trong cơ sở dữ liệu các mã ngành thuộc nhóm Cấp 1 tìm thấy, hoặc toàn bộ nếu không tìm thấy nhóm Cấp 1 cụ thể
  const targetCodes = Object.keys(vsicRawData).filter(code => {
    if (bestMainSector) {
      // Chỉ tìm các mã có cha phân cấp dẫn tới bestMainSector
      const hier = getSectorHierarchy(code);
      return hier["1"]?.ma === bestMainSector;
    }
    return true;
  });

  let bestMatch: { ma: string; ten: string; diem: number } | null = null;
  let highestScore = 0;

  for (const ma of targetCodes) {
    const ten = vsicRawData[ma];
    const tenLower = ten.toLowerCase();
    
    // Tính điểm tương đồng dựa trên từ khóa chung
    let score = 0;
    const wordsInTen = tenLower.split(/[\s,.\-/]+/);
    const wordsInDesc = text.split(/[\s,.\-/]+/);

    // Điểm trùng khớp cụm từ đầy đủ (phạt theo độ rộng)
    if (text.includes(tenLower)) {
      score += 10.0 + (tenLower.length / text.length) * 5;
    }

    // Điểm trùng khớp từng từ
    let matchedWords = 0;
    for (const w of wordsInTen) {
      if (w.length > 2 && text.includes(w)) {
        matchedWords++;
      }
    }
    score += (matchedWords / wordsInTen.length) * 5.0;

    // Ưu tiên các mã ngành cụ thể (cấp 4, cấp 5) nếu mô tả rất khớp
    if (ma.length >= 4) {
      score += 1.2;
    }

    if (score > highestScore) {
      highestScore = score;
      bestMatch = {
        ma,
        ten,
        diem: Math.min(score / 15.0, 1.0) // Chuẩn hóa điểm số từ 0.0 đến 1.0
      };
    }
  }

  // Fallback: nếu không tìm thấy gì nổi bật, tìm trên toàn bộ danh mục
  if (!bestMatch || bestMatch.diem < 0.25) {
    let fallbackMatch: { ma: string; ten: string; diem: number } | null = null;
    let fallbackHighest = 0;
    
    for (const ma of Object.keys(vsicRawData)) {
      const ten = vsicRawData[ma];
      if (text.includes(ma.toLowerCase())) {
        fallbackMatch = { ma, ten, diem: 0.9 };
        break;
      }
      
      const tenLower = ten.toLowerCase();
      let matchedCount = 0;
      const keywords = tenLower.split(" ");
      for (const kw of keywords) {
        if (kw.length > 2 && text.includes(kw)) {
          matchedCount++;
        }
      }
      const score = (matchedCount / keywords.length);
      if (score > fallbackHighest) {
        fallbackHighest = score;
        fallbackMatch = { ma, ten, diem: score };
      }
    }
    
    if (fallbackMatch && fallbackMatch.diem > 0.3) {
      return fallbackMatch;
    }
  }

  return bestMatch;
}
