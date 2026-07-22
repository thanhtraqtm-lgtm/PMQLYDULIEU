export interface PipelineStep {
  id: number;
  title: string;
  shortDesc: string;
  fullTitle: string;
  fullDesc: string;
  highlights: string[];
  actionText: string;
  targetTab: string;
}

export interface ActionButton {
  text: string;
  tab: string;
  stepId?: number;
  style?: "primary" | "secondary";
}

export interface GuideScenario {
  id: string;
  icon: string;
  buttonText: string;
  title: string;
  intro: string;
  steps: string[];
  actionButtons: ActionButton[];
}

export const PIPELINE_STEPS: PipelineStep[] = [
  {
    id: 1,
    title: "Nạp & Định nghĩa",
    shortDesc: "Tải file Excel/CSV thô, đặt lại tên cột và lọc các cột cần thiết.",
    fullTitle: "Nạp & Định nghĩa Cột (Khởi đầu bắt buộc)",
    fullDesc: "Mọi tệp Excel khảo sát từ các địa bàn gửi về có thể có cấu trúc tiêu đề cột khác nhau (ví dụ: Ma_DN, MADT, Mã đơn vị). Bạn cần đưa vào hệ thống và gán nhãn chuẩn hóa để các tính năng phân tích phía sau tự động nhận diện chính xác.",
    highlights: [
      "Tệp đầu vào: Excel (.xlsx) hoặc CSV thô",
      "Đặt lại tên cột dễ nhớ, loại bỏ cột thừa"
    ],
    actionText: "📂 NẠP FILE & ĐỊNH NGHĨA CỘT NGAY",
    targetTab: "xemdulieu"
  },
  {
    id: 2,
    title: "Kiểm soát & Logic",
    shortDesc: "Đặt quy tắc logic, quét phát hiện mâu thuẫn doanh thu & lao động.",
    fullTitle: "Rà soát Quy tắc Logic & Phát hiện Lệch quy luật",
    fullDesc: "Tránh các lỗi nhập liệu phi lý trong dữ liệu khảo sát. Bạn có thể tự viết các biểu thức logic dễ hiểu trong Rules Studio (ví dụ: DoanhThu > 0 => LaoDong > 0). Hệ thống sẽ chạy quét tức thì, bôi đỏ các ô bị vi phạm và hiển thị danh sách cảnh báo chi tiết.",
    highlights: [
      "Tự thiết lập quy tắc rà soát linh hoạt",
      "Phát hiện Outliers (giá trị đột biến, bất thường)"
    ],
    actionText: "🛡️ VÀO RULES STUDIO (QUY TẮC)",
    targetTab: "rulesstudio"
  },
  {
    id: 3,
    title: "Tổng hợp & Xuất",
    shortDesc: "Lập bảng đa chiều Pivot, tính tần suất và tải báo cáo sạch Excel.",
    fullTitle: "Tổng hợp Báo cáo đa chiều & Xuất Excel Sạch",
    fullDesc: "Sau khi dữ liệu đã được nạp và lọc sạch lỗi logic, đây là lúc tạo ra sản phẩm cuối cùng. Thiết lập các trục dòng (Row), trục cột (Column) và chỉ tiêu tính toán (Doanh thu, Lao động, Số đơn vị) để hệ thống tự động vẽ bảng Pivot đa chiều và kết xuất file Excel chuẩn báo cáo.",
    highlights: [
      "Lập bảng thống kê Pivot nhanh trong 3 giây",
      "Xuất file báo cáo sạch không tì vết"
    ],
    actionText: "📈 ĐẾN TRANG TỔNG HỢP BÁO CÁO",
    targetTab: "tonghop"
  }
];

export const GUIDE_SCENARIOS: GuideScenario[] = [
  {
    id: "naptiep",
    icon: "📥",
    buttonText: "Tôi có file Excel thô mới",
    title: "Lộ trình: Xử lý tệp dữ liệu thô (Excel / CSV mới nạp)",
    intro: "Khi bạn nhận được một tệp khảo sát mới từ bên ngoài, cấu trúc cột có thể lộn xộn. Hãy đi theo quy trình 3 bước cực kỳ đơn giản sau:",
    steps: [
      "Vào tab “Quản lý Tệp” ➔ “Xem & Định nghĩa cột”. Tải tệp lên.",
      "Tìm dòng “Định nghĩa / Ánh xạ cột dữ liệu” phía dưới. Tại đây, hãy chỉ ra đâu là cột chứa Mã đơn vị, Doanh thu, Số lao động, hay Tỉnh thành bằng cách chọn nhãn tương ứng ở hộp thả xuống.",
      "Nếu có cột viết tắt khó hiểu, bạn có thể nhấp đúp để đặt lại tên tiếng Việt trực quan. Sau khi gán xong, dữ liệu lập tức hiển thị sạch sẽ ở bảng xem trước!"
    ],
    actionButtons: [
      {
        text: "👉 Đến trang nạp tệp & làm theo ngay",
        tab: "xemdulieu",
        stepId: 1,
        style: "primary"
      }
    ]
  },
  {
    id: "chuanhoamanganh",
    icon: "🏷️",
    buttonText: "Tôi muốn rà soát lệch mã ngành VSIC",
    title: "Lộ trình: Đối chiếu mâu thuẫn giữa Mô tả thực tế và Mã ngành đã chọn",
    intro: "Trong thực tế điều tra, mã ngành do điều tra viên chọn trên máy chắc chắn hợp lệ về mặt cú pháp (vì hệ thống buộc phải chọn từ danh mục chuẩn). Tuy nhiên, lỗi rất dễ xảy ra ở việc áp sai mã so với mô tả thực tế, hoặc mô tả một kiểu nhưng mã chọn một đằng. Hãy xử lý theo quy trình rà soát:",
    steps: [
      "Vào tab “Trí tuệ VSIC” ➔ “Kiểm tra mã ngành VSIC”. Chọn cột Mô tả hoạt động và cột Mã ngành đã áp. Hệ thống sẽ quét chéo và phát hiện các trường hợp mâu thuẫn (Ví dụ: Mô tả ghi 'Trồng lúa' nhưng mã lại chọn 'Nuôi cá').",
      "Xem chi tiết lỗi để điều phối điều tra viên phúc tra trực tiếp nhằm đảm bảo tính thống nhất dữ liệu."
    ],
    actionButtons: [
      {
        text: "🤖 Rà soát mâu thuẫn mã & mô tả",
        tab: "chuanhoanganh",
        stepId: 2,
        style: "primary"
      }
    ]
  },
  {
    id: "soatlogic",
    icon: "🕵️",
    buttonText: "Tôi cần tìm lỗi sai logic dữ liệu",
    title: "Lộ trình: Chạy rà soát lỗi logic dữ liệu & Outliers",
    intro: "Để đảm bảo dữ liệu trước khi xuất báo cáo không có lỗi phi lý (ví dụ: Đơn vị đang hoạt động nhưng doanh thu bằng 0, số lao động âm, hay lương trung bình cao bất thường):",
    steps: [
      "Vào tab “Kiểm soát & Logic” ➔ “Rules Studio”. Ở đây đã có sẵn một loạt quy tắc rà soát mẫu (ví dụ kiểm tra Doanh thu và Lao động). Bạn có thể kích hoạt các quy tắc này hoặc tự gõ biểu thức mới cực kỳ đơn giản.",
      "Vào tab phụ “Kiểm tra Logic”, chọn chạy kiểm tra. Hệ thống sẽ quét toàn bộ bảng và chỉ rõ bao nhiêu dòng vi phạm quy tắc nào.",
      "Vào tab “Quét lệch quy luật” (Outliers) để hệ thống tự động vẽ biểu đồ phân tán và tìm ra các doanh nghiệp có mức doanh thu/lao động lệch hẳn so với mặt bằng chung (có thể do gõ thừa số 0)."
    ],
    actionButtons: [
      {
        text: "🛡️ Thiết lập Quy tắc rà soát",
        tab: "rulesstudio",
        stepId: 3,
        style: "primary"
      },
      {
        text: "📈 Quét giá trị bất thường (Outliers)",
        tab: "outliers",
        stepId: 3,
        style: "secondary"
      }
    ]
  },
  {
    id: "gopfile",
    icon: "🔗",
    buttonText: "Tôi muốn gộp 2 file Excel riêng biệt",
    title: "Lộ trình: Ghép nối (Join) hai bảng Excel riêng biệt thành một",
    intro: "Nếu bạn có bảng A (chứa danh sách doanh nghiệp và doanh thu) và bảng B (chứa mã số doanh nghiệp và danh sách địa bàn quận huyện), bạn muốn gộp chúng lại dựa trên Mã số chung:",
    steps: [
      "Vào tab “Quản lý Tệp” ➔ “Ghép nối dữ liệu”.",
      "Tải bảng chính lên làm Tệp dữ liệu gốc, tải bảng phụ chứa thông tin bổ sung làm Tệp dữ liệu bổ sung.",
      "Chọn Cột liên kết chung (ví dụ: cột mã số thuế / mã định danh) ở cả hai bảng. Chọn cột ở bảng phụ mà bạn muốn thêm vào bảng chính (ví dụ: cột Địa bàn).",
      "Nhấn “Thực hiện ghép nối”. Hệ thống sẽ tự động ghép Left Join hoàn hảo và cho phép bạn tải ngay file Excel kết quả đã gộp về máy tính!"
    ],
    actionButtons: [
      {
        text: "🔗 Bắt đầu ghép nối 2 file ngay",
        tab: "ghepnoi",
        stepId: 1,
        style: "primary"
      }
    ]
  },
  {
    id: "baocao",
    icon: "📊",
    buttonText: "Tôi muốn vẽ biểu đồ & lập Pivot",
    title: "Lộ trình: Tạo bảng Pivot đa chiều, Tần suất & Vẽ biểu đồ",
    intro: "Khi dữ liệu đã hoàn toàn sạch sẽ, việc kết xuất báo cáo đa chiều được thực hiện tự động chỉ bằng vài cú nhấp chuột:",
    steps: [
      "Vào tab “Tổng hợp Báo cáo” ➔ “Tổng hợp báo cáo đa chiều”.",
      "Chọn Trục dòng (ví dụ: phân nhóm theo Tỉnh thành hoặc Mã ngành cấp 1/2), chọn Trục cột (nếu có), và chọn Chỉ tiêu tính toán (Ví dụ: Tổng doanh thu, Số lao động trung bình).",
      "Bảng Pivot đa chiều sẽ tự động dựng và tính toán tức thì. Bạn có thể xem trực tiếp, xem tỷ lệ phần trăm hoặc chuyển sang tab Tần suất xuất hiện / Phân tích tương quan để vẽ biểu đồ và tải báo cáo sạch!"
    ],
    actionButtons: [
      {
        text: "📈 Thiết lập bảng Pivot & Vẽ biểu đồ ngay",
        tab: "tonghop",
        stepId: 4,
        style: "primary"
      }
    ]
  },
  {
    id: "sampling",
    icon: "🎲",
    buttonText: "Tôi muốn lấy mẫu thanh tra (Sampling)",
    title: "Lộ trình: Thiết lập Lấy mẫu thanh tra phân tầng (Sampling)",
    intro: "Khi số lượng doanh nghiệp khảo sát quá lớn (ví dụ: hàng vạn doanh nghiệp), việc kiểm tra thủ công toàn bộ là bất khả thi. Bạn cần lấy một tập mẫu đại diện nhưng vẫn đảm bảo tính khách quan và khoa học để thanh tra trực tiếp:",
    steps: [
      "Vào tab “Quản lý Tệp” ➔ “Chọn mẫu ngẫu nhiên” (hoặc nhấp trực tiếp vào nút hành động bên dưới).",
      "Xác định phương pháp lấy mẫu: “Ngẫu nhiên đơn giản” (Simple Random) hoặc “Phân tầng tỉ lệ” (Stratified Sampling) theo ngành nghề/vị trí địa lý để đảm bảo nhóm nào cũng có đại diện phù hợp.",
      "Chọn kích thước mẫu (ví dụ: lấy 5% tổng số dòng hoặc cố định 100 doanh nghiệp), sau đó hệ thống sẽ tự động rút trích và bôi đậm danh sách được chọn. Bạn có thể tải ngay tệp đã lấy mẫu về để thực thi!"
    ],
    actionButtons: [
      {
        text: "🎲 Đến công cụ Lấy mẫu thanh tra",
        tab: "chonmau",
        stepId: 1,
        style: "primary"
      }
    ]
  },
  {
    id: "yoy",
    icon: "🔄",
    buttonText: "Tôi muốn đối sánh dữ liệu liên năm (YoY)",
    title: "Lộ trình: Đối sánh hiệu số giữa các năm & Phát hiện tăng trưởng đột biến",
    intro: "Để phân tích tốc độ phát triển kinh tế xã hội và so sánh xem doanh nghiệp, địa bàn nào có doanh thu/lao động sụt giảm hoặc tăng vọt so với kỳ trước:",
    steps: [
      "Vào tab “Kiểm soát & Logic” ➔ “Đối sánh dữ liệu”. Tại đây, hãy tải lên biểu tổng hợp của Năm trước (Năm A) và Năm nay (Năm B).",
      "Chọn Cột khóa liên kết chung (ví dụ: Địa bàn hoặc Mã Số Thuế) và thiết lập các cột muốn đối chiếu (Doanh thu, Lao động).",
      "Hệ thống sẽ tự động ghép nối và tính ra “Tốc độ tăng trưởng %” cùng với “Giá trị tuyệt đối chênh lệch”. Bảng báo cáo phân tích biến động sẽ bôi màu xanh lá nếu tăng trưởng tốt và màu đỏ nếu sụt giảm bất thường!"
    ],
    actionButtons: [
      {
        text: "🔄 Đến công cụ Đối sánh YoY",
        tab: "sosanh",
        stepId: 3,
        style: "primary"
      }
    ]
  },
  {
    id: "collab",
    icon: "👥",
    buttonText: "Tôi muốn cộng tác nhập liệu & Quản trị đám mây",
    title: "Lộ trình: Nhập liệu đồng bộ đám mây & Cộng tác liên ngành",
    intro: "Thay vì gửi file Excel qua lại dễ thất lạc và nhầm lẫn phiên bản, bạn có thể đồng bộ hóa dữ liệu khảo sát trực tiếp lên máy chủ đám mây của Chính phủ:",
    steps: [
      "Vào tab “Cộng tác liên ngành” ➔ “Nhập liệu trực tuyến”. Tại đây, các thành viên thuộc các tỉnh thành khác nhau có thể đăng nhập bằng email đơn vị và cùng nhau nạp dữ liệu lên một bảng tổng hợp chung thời gian thực.",
      "Quản trị viên có thể vào “Quản trị hệ thống” (Admin Dashboard) để duyệt tài khoản, phân quyền thao tác và theo dõi hoạt động nạp dữ liệu của từng địa bàn."
    ],
    actionButtons: [
      {
        text: "📝 Vào form nhập liệu đám mây",
        tab: "dataentry",
        stepId: 1,
        style: "primary"
      },
      {
        text: "🛡️ Vào Quản trị hệ thống",
        tab: "admindashboard",
        stepId: 1,
        style: "secondary"
      }
    ]
  },
  {
    id: "excelsql",
    icon: "🤖",
    buttonText: "Tôi muốn viết công thức Excel & truy vấn SQL bằng AI",
    title: "Lộ trình: Trợ lý AI tạo tự động công thức Excel & câu lệnh SQL",
    intro: "Khi xử lý bảng dữ liệu lớn, việc nhớ chính xác cú pháp hàm Excel phức tạp hay câu lệnh SQL tốn nhiều thời gian. Hãy để Trợ lý AI hỗ trợ bạn:",
    steps: [
      "Vào tab “Tiện ích & Cộng tác” ➔ “Trợ lý Excel & Truy vấn SQL AI”.",
      "Mô tả yêu cầu bằng ngôn ngữ tự nhiên (Ví dụ: 'Tính tổng doanh thu theo tỉnh thành', 'Viết câu lệnh SQL lọc doanh nghiệp có trên 50 lao động').",
      "Trợ lý AI sẽ phân tích cấu trúc dữ liệu đang mở và sinh ra công thức/truy vấn chuẩn xác kèm giải thích chi tiết để bạn sử dụng ngay."
    ],
    actionButtons: [
      {
        text: "🤖 Mở Trợ lý Excel & SQL AI",
        tab: "excelsql",
        stepId: 1,
        style: "primary"
      }
    ]
  },
  {
    id: "outliers",
    icon: "🚨",
    buttonText: "Tôi muốn phát hiện doanh nghiệp bất thường (Outliers)",
    title: "Lộ trình: Phát hiện bất thường & Đột biến số liệu (Statistical Outliers)",
    intro: "Khi có hàng ngàn bản ghi khảo sát, việc các doanh nghiệp khai báo khống, ghi nhầm thêm chữ số 0, hoặc có số liệu bất thường về doanh thu/thu nhập bình quân là rất phổ biến. Thuật toán phân tích ngoại lai (Outliers) giúp lọc nhanh ra các nghi vấn để rà soát lại:",
    steps: [
      "Vào tab “Kiểm soát & Logic” ➔ “Phát hiện ngoại lai”. Chọn cột chỉ số cần kiểm tra (Ví dụ: Doanh thu, Lao động, Thu nhập).",
      "Chọn phương pháp phân tích: “Z-Score” (Phù hợp phân phối chuẩn) hoặc “IQR / Boxplot” (Phù hợp với dữ liệu không đồng đều) và thiết lập hệ số nhạy cảm.",
      "Hệ thống sẽ tự động vẽ biểu đồ Boxplot trực quan và lọc ra danh sách các doanh nghiệp nằm ngoài dải an toàn để điều tra viên phúc tra trực tiếp."
    ],
    actionButtons: [
      {
        text: "🚨 Phân tích ngoại lai ngay",
        tab: "outliers",
        stepId: 1,
        style: "primary"
      }
    ]
  },
  {
    id: "splitfile",
    icon: "✂️",
    buttonText: "Tôi muốn chia tách tệp Excel lớn thành nhiều tệp con",
    title: "Lộ trình: Phân tách tệp tổng hợp Excel theo địa bàn/nhóm ngành",
    intro: "Khi nhận được một tệp Excel tổng hợp quy mô cả nước hoặc toàn tỉnh, bạn cần chia nhỏ dữ liệu này thành từng tệp riêng biệt để gửi cho các quận, huyện hoặc các tổ điều tra độc lập xử lý:",
    steps: [
      "Vào tab “Quản lý Tệp” ➔ “Tách tệp Excel”. Tải lên tệp tổng hợp mà bạn muốn chia nhỏ.",
      "Chọn Cột phân loại chính để tách (Ví dụ: Cột “Tỉnh/Thành phố”, “Quận/Huyện” hoặc “Mã ngành cấp 1”).",
      "Hệ thống sẽ tự động quét toàn bộ tệp, nhóm dữ liệu và tạo ra danh sách các tệp con tương ứng. Bạn chỉ cần nhấp một nút duy nhất để tải xuống toàn bộ các tệp đã phân tách dưới dạng file nén ZIP!"
    ],
    actionButtons: [
      {
        text: "✂️ Đến công cụ Tách tệp Excel",
        tab: "tachfile",
        stepId: 1,
        style: "primary"
      }
    ]
  }
];
