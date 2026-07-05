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
    title: "Chuẩn hóa VSIC (AI)",
    shortDesc: "Đối chiếu mã ngành, dùng Trí tuệ Nhân tạo dịch ngôn ngữ sang mã.",
    fullTitle: "Chuẩn hóa VSIC & Trí tuệ Nhân tạo AI",
    fullDesc: "Mã ngành đăng ký kinh doanh thường bị sai sót hoặc chỉ ghi chữ mô tả tiếng Việt không có mã số chuẩn. Công cụ này đối sánh với danh mục VSIC chuẩn quốc gia. Đặc biệt, mô hình AI tích hợp sẽ tự động phân tích câu từ tiếng Việt và dịch sang mã số chính xác (Ví dụ: 'Nuôi cá nước ngọt' ➔ 03210).",
    highlights: [
      "Quét mã sai, tìm ngành phù hợp",
      "Sử dụng AI tự động dịch mô tả sang mã số"
    ],
    actionText: "🤖 SỬ DỤNG AI ĐỐI SÁNH",
    targetTab: "smartcatalog"
  },
  {
    id: 3,
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
    id: 4,
    title: "Tổng hợp & Xuất",
    shortDesc: "Lập bảng đa chiều Pivot, tính tần suất và tải báo cáo sạch Excel.",
    fullTitle: "Tổng hợp Báo cáo đa chiều & Xuất Excel Sạch",
    fullDesc: "Sau khi dữ liệu đã được nạp, chuẩn hóa mã ngành và lọc sạch lỗi logic, đây là lúc tạo ra sản phẩm cuối cùng. Thiết lập các trục dòng (Row), trục cột (Column) và chỉ tiêu tính toán (Doanh thu, Lao động, Số đơn vị) để hệ thống tự động vẽ bảng Pivot đa chiều và kết xuất file Excel chuẩn báo cáo.",
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
    buttonText: "Tôi muốn sửa lỗi mã ngành VSIC",
    title: "Lộ trình: Chuẩn hóa & Sửa lỗi mã ngành VSIC",
    intro: "Có 2 trường hợp phổ biến khi chuẩn hóa mã ngành cho doanh nghiệp trong danh sách khảo sát của bạn:",
    steps: [
      "Trường hợp A (Có sẵn mã nhưng sợ sai): Vào tab “Trí tuệ VSIC” ➔ “Kiểm tra mã ngành VSIC”. Chọn cột chứa mã ngành trong file của bạn. Hệ thống sẽ tự kiểm tra xem mã đó có khớp với bảng VSIC chuẩn hay không và đánh dấu đỏ các dòng sai.",
      "Trường hợp B (Chỉ ghi chữ tiếng Việt, không có mã số): Vào tab “Trí tuệ VSIC” ➔ “Đối sánh danh mục (AI)”. Chọn cột chứa nội dung tiếng Việt mô tả công việc (ví dụ: 'Sản xuất phần mềm'). Trí tuệ nhân tạo AI (Gemini) sẽ quét tự động và gợi ý mã số VSIC cấp 5 chính xác cho bạn, tiết kiệm hàng tuần tra cứu thủ công!"
    ],
    actionButtons: [
      {
        text: "🤖 Sử dụng AI gán mã ngành tự động",
        tab: "smartcatalog",
        stepId: 2,
        style: "primary"
      },
      {
        text: "🔍 Kiểm tra mã ngành đã có sẵn",
        tab: "chuanhoanganh",
        stepId: 2,
        style: "secondary"
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
  }
];
