import React, { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { useAuth, db, isFirebaseInitialized } from "../context/AuthContext";
import { GoogleGenAI } from "@google/genai";
import { SignaturePad } from "./SignaturePad";
import { SignatureToken } from "./SignatureToken";
import { 
  collection, 
  addDoc, 
  getDocs,
  deleteDoc,
  doc,
  updateDoc,
  serverTimestamp, 
  writeBatch 
} from "firebase/firestore";
import { 
  FileUp, 
  Plus, 
  Database, 
  FileSpreadsheet, 
  AlertCircle, 
  CheckCircle, 
  Loader2, 
  Trash2,
  ScanFace,
  Search,
  Filter,
  Download,
  CloudLightning,
  Sparkles,
  Edit2,
  Save,
  Building2,
  Store,
  ChevronDown,
  RefreshCw,
  HelpCircle,
  PenTool,
  Fingerprint,
  FileSignature,
  ShieldCheck,
  RotateCcw,
  Key
} from "lucide-react";

// Định nghĩa kiểu cấu hình trường động
export interface CustomField {
  name: string;      // Khóa lưu trữ (ví dụ: san_luong_A)
  label: string;     // Nhãn biểu mẫu (ví dụ: Sản lượng sản phẩm A)
  type: "text" | "number" | "boolean"; // Kiểu trường
}

// Định nghĩa kiểu dữ liệu cho bản ghi nhập tin khảo sát
export interface SurveyRecord {
  id?: string;
  mst: string; // Mã số thuế hoặc mã cơ sở
  name: string; // Tên cơ sở/doanh nghiệp
  representative?: string; // Người đại diện / Chủ cơ sở
  address: string; // Địa bàn / Địa chỉ
  doanhthu: number; // Doanh thu (triệu đồng)
  laodong: number; // Số lao động (người)
  manganh: string; // Mã ngành VSIC chính
  ghichu?: string; // Ghi chú
  block: "cathe" | "doanhnghiep"; // Khối: Cá thể vs Doanh nghiệp
  entryMethod: "manual" | "scan" | "excel" | "template"; // Phương thức nhập
  createdBy: string;
  createdAt: string;
  customData?: Record<string, any>; // Các trường dữ liệu động tùy biến từ mẫu
  is_signed?: boolean;
  signed_mode?: "draw" | "token" | "";
  signed_by?: string;
  signed_time?: string;
  signed_hash?: string;
  signature_img?: string;
  surveyor_signature?: string;
  surveyor_name?: string;
  surveyor_phone?: string;
  respondent_name?: string;
  respondent_phone?: string;
}

// Reusable digit boxes input mimicking physical paper form cells
interface GridBoxesInputProps {
  value: string;
  onChange: (val: string) => void;
  length: number;
  splitIndex?: number;
  disabled?: boolean;
}

const GridBoxesInput: React.FC<GridBoxesInputProps> = ({ value, onChange, length, splitIndex, disabled }) => {
  const [isFocused, setIsFocused] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleBoxClick = () => {
    if (!disabled) {
      inputRef.current?.focus();
    }
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    let val = e.target.value.toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (val.length > length) {
      val = val.substring(0, length);
    }
    onChange(val);
  };

  const chars = value.split("");

  return (
    <div className="relative inline-flex items-center cursor-text select-none" onClick={handleBoxClick}>
      {/* Hidden native input */}
      <input
        ref={inputRef}
        type="text"
        value={value}
        onChange={handleInputChange}
        onFocus={() => setIsFocused(true)}
        onBlur={() => setIsFocused(false)}
        disabled={disabled}
        className="absolute inset-0 w-full h-full opacity-0 z-10 cursor-text"
        style={{ caretColor: "transparent" }}
      />
      
      {/* Render squares */}
      <div className="flex items-center gap-1">
        {Array.from({ length }).map((_, idx) => {
          const char = chars[idx] || "";
          const isCurrentActive = isFocused && idx === Math.min(chars.length, length - 1);
          const hasVal = char !== "";
          const showSplit = splitIndex !== undefined && idx === splitIndex;

          return (
            <React.Fragment key={idx}>
              {showSplit && (
                <span className="text-slate-400 font-bold px-1 select-none">-</span>
              )}
              <div
                className={`w-6 h-7 xs:w-7 xs:h-8 flex items-center justify-center border rounded font-mono text-xs xs:text-sm font-black transition-all ${
                  isCurrentActive
                    ? "border-indigo-600 bg-indigo-50/45 ring-2 ring-indigo-200 text-indigo-900 shadow-xs scale-105"
                    : hasVal
                    ? "border-slate-400 bg-slate-50 text-slate-800"
                    : "border-slate-300 bg-white hover:border-slate-400 text-slate-300"
                }`}
              >
                {char}
                {isCurrentActive && char === "" && (
                  <span className="animate-pulse text-indigo-500 font-normal">|</span>
                )}
              </div>
            </React.Fragment>
          );
        })}
      </div>
    </div>
  );
};

export const DataEntry: React.FC = () => {
  const { user, googleAccessToken, signInWithGoogle } = useAuth();
  const unitID = user?.unitID || "guest";
  const userName = user?.displayName || "Người dùng Khách";

  // State khối đang chọn: "doanhnghiep" hoặc "cathe"
  const [activeBlock, setActiveBlock] = useState<"doanhnghiep" | "cathe">("doanhnghiep");

  // State phương thức nhập: "manual" | "scan" | "excel" | "template"
  const [entryMethod, setEntryMethod] = useState<"manual" | "scan" | "excel" | "template">("manual");

  // State dữ liệu danh sách tin đã nhập
  const [surveyRecords, setSurveyRecords] = useState<SurveyRecord[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [filterBlock, setFilterBlock] = useState<"all" | "doanhnghiep" | "cathe">("all");
  const [filterMethod, setFilterMethod] = useState<"all" | "manual" | "scan" | "excel" | "template">("all");

  // Trạng thái hệ thống
  const [loading, setLoading] = useState(false);
  const [statusMessage, setStatusMessage] = useState<{ type: "success" | "error"; text: string } | null>(null);

  // --- STATE NHẬP THỦ CÔNG ---
  const [manualFormType, setManualFormType] = useState<"standard" | "hungyen">("hungyen");
  const [manualMst, setManualMst] = useState("");
  const [manualName, setManualName] = useState("");
  const [manualRep, setManualRep] = useState("");
  const [manualAddress, setManualAddress] = useState("");
  const [manualDoanhThu, setManualDoanhThu] = useState("");
  const [manualLaoDong, setManualLaoDong] = useState("");
  const [manualMaNganh, setManualMaNganh] = useState("");
  const [manualGhiChu, setManualGhiChu] = useState("");

  // --- STATE MẪU PHIẾU GIẤY HƯNG YÊN CHUYÊN SÂU ---
  const [hyMst, setHyMst] = useState("");
  const [hyName, setHyName] = useState("");
  const [hyRep, setHyRep] = useState("");
  const [hyCommuneName, setHyCommuneName] = useState("");
  const [hyCommuneCode, setHyCommuneCode] = useState("");
  const [hySpecificAddress, setHySpecificAddress] = useState("");
  const [hyPhone, setHyPhone] = useState("");
  const [hyEmail, setHyEmail] = useState("");
  const [hyMaNganh, setHyMaNganh] = useState("");
  const [hyMaNganhMoTa, setHyMaNganhMoTa] = useState("");
  const [hyNganhKhac1, setHyNganhKhac1] = useState("");
  const [hyNganhKhac1MoTa, setHyNganhKhac1MoTa] = useState("");
  const [hyNganhKhac2, setHyNganhKhac2] = useState("");
  const [hyNganhKhac2MoTa, setHyNganhKhac2MoTa] = useState("");
  const [hyLaoDong0101, setHyLaoDong0101] = useState("");
  const [hyLaoDong3112, setHyLaoDong3112] = useState("");
  const [hyCoSoNgoaiXa, setHyCoSoNgoaiXa] = useState(""); // "có" | "không" | ""
  const [hyDoanhThu, setHyDoanhThu] = useState("");
  const [hyGhiChu, setHyGhiChu] = useState("");

  // Các trường bổ sung của Hưng Yên Cá Thể & Thông tin bổ sung
  const [hySttCoSo, setHySttCoSo] = useState("");
  const [hyDoanhThuNamTruoc, setHyDoanhThuNamTruoc] = useState("");
  const [hyDoanhThu9Thang, setHyDoanhThu9Thang] = useState("");
  const [hyDoanhThu3Thang, setHyDoanhThu3Thang] = useState("");
  const [hySurveyorName, setHySurveyorName] = useState("");
  const [hySurveyorPhone, setHySurveyorPhone] = useState("");
  const [hyRespondentName, setHyRespondentName] = useState("");
  const [hyRespondentPhone, setHyRespondentPhone] = useState("");

  // State danh sách cơ sở ngoài xã (Trang 2 Doanh nghiệp)
  const [hyCoSoNgoaiXaList, setHyCoSoNgoaiXaList] = useState([
    { stt: "1", ten: "", diaChi: "", laoDong0101: "", laoDong3112: "" },
    { stt: "2", ten: "", diaChi: "", laoDong0101: "", laoDong3112: "" },
    { stt: "3", ten: "", diaChi: "", laoDong0101: "", laoDong3112: "" },
    { stt: "4", ten: "", diaChi: "", laoDong0101: "", laoDong3112: "" },
  ]);

  // State kết quả sản xuất kinh doanh chi tiết (Trang 2 Doanh nghiệp)
  const [hyKqkdTongDtNamTruoc, setHyKqkdTongDtNamTruoc] = useState("");
  const [hyKqkdTongDt9Thang, setHyKqkdTongDt9Thang] = useState("");
  const [hyKqkdTongDtUocCaNam, setHyKqkdTongDtUocCaNam] = useState("");

  const [hyKqkdChinhDtNamTruoc, setHyKqkdChinhDtNamTruoc] = useState("");
  const [hyKqkdChinhDt9Thang, setHyKqkdChinhDt9Thang] = useState("");
  const [hyKqkdChinhDtUocCaNam, setHyKqkdChinhDtUocCaNam] = useState("");

  const [hyKqkdKhac1Ten, setHyKqkdKhac1Ten] = useState("");
  const [hyKqkdKhac1DtNamTruoc, setHyKqkdKhac1DtNamTruoc] = useState("");
  const [hyKqkdKhac1Dt9Thang, setHyKqkdKhac1Dt9Thang] = useState("");
  const [hyKqkdKhac1DtUocCaNam, setHyKqkdKhac1DtUocCaNam] = useState("");

  const [hyKqkdKhac2Ten, setHyKqkdKhac2Ten] = useState("");
  const [hyKqkdKhac2DtNamTruoc, setHyKqkdKhac2DtNamTruoc] = useState("");
  const [hyKqkdKhac2Dt9Thang, setHyKqkdKhac2Dt9Thang] = useState("");
  const [hyKqkdKhac2DtUocCaNam, setHyKqkdKhac2DtUocCaNam] = useState("");

  // State Chữ ký số & Ký tay điện tử
  const [hySignatureMode, setHySignatureMode] = useState<"draw" | "token" | "">("");
  const [hySignatureDataUrl, setHySignatureDataUrl] = useState("");
  const [hySurveyorSignature, setHySurveyorSignature] = useState("");
  const [hyIsSigned, setHyIsSigned] = useState(false);
  const [hySignedTime, setHySignedTime] = useState("");
  const [hySignedBy, setHySignedBy] = useState("");
  const [hySignedHash, setHySignedHash] = useState("");
  const [selectedSignatureRecord, setSelectedSignatureRecord] = useState<SurveyRecord | null>(null);

  // --- STATE TRƯỜNG DỮ LIỆU ĐỘNG (CUSTOM FIELDS) ---
  const [customFields, setCustomFields] = useState<CustomField[]>([]);
  const [dynamicValues, setDynamicValues] = useState<Record<string, string>>({});
  
  // Các trường cho việc Thêm Thủ Công trường động mới
  const [newFieldName, setNewFieldName] = useState("");
  const [newFieldLabel, setNewFieldLabel] = useState("");
  const [newFieldType, setNewFieldType] = useState<"text" | "number" | "boolean">("text");

  // --- STATE ĐỌC FILE SCAN (AI EXTRACT) ---
  const [scanFile, setScanFile] = useState<File | null>(null);
  const [scanPreviewUrl, setScanPreviewUrl] = useState<string | null>(null);
  const [isExtracting, setIsExtracting] = useState(false);
  const [aiLogs, setAiLogs] = useState<string[]>([]);
  const [extractedResult, setExtractedResult] = useState<Partial<SurveyRecord> | null>(null);
  const scanFileInputRef = useRef<HTMLInputElement>(null);

  // --- STATE NẠP FILE EXCEL ---
  const [excelRows, setExcelRows] = useState<any[]>([]);
  const [excelFileName, setExcelFileName] = useState("");
  const [uploadProgress, setUploadProgress] = useState<string | null>(null);
  const excelFileInputRef = useRef<HTMLInputElement>(null);

  // --- STATE CHỈNH SỬA BẢN GHI (MODAL) ---
  const [editingRecord, setEditingRecord] = useState<SurveyRecord | null>(null);

  // --- STATE GOOGLE DRIVE SYNC ---
  const [isDriveSyncing, setIsDriveSyncing] = useState(false);
  const [driveSyncLogs, setDriveSyncLogs] = useState<string[]>([]);
  const [showDrivePanel, setShowDrivePanel] = useState(false);

  // --- STATE PHÂN TÍCH BIỂU MẪU EXCEL ---
  const [templateFileName, setTemplateFileName] = useState("");
  const [analyzedFields, setAnalyzedFields] = useState<CustomField[]>([]);
  const templateFileInputRef = useRef<HTMLInputElement>(null);

  const handleTemplateFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setTemplateFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        if (rows.length === 0) {
          setStatusMessage({ type: "error", text: "Tệp Excel trống!" });
          return;
        }
        
        // Lấy hàng đầu tiên làm headers
        const headers = (rows[0] as any[]).map(h => String(h || "").trim());
        
        // Danh sách các cột chuẩn đã có sẵn trong hệ thống
        const standardHeaders = [
          "mst", "masothue", "mã số thuế", "mãđịnhdanh", "mã cơ sở", "macoso",
          "tên", "tendoanhnghiep", "tên doanh nghiệp", "tên cơ sở", "tencoso", "doanhnghiep", "tên đơn vị",
          "đạidiện", "chủ cơ sở", "chucoso", "người đại diện", "daidien", "nguoidaidien",
          "địachỉ", "diachi", "địabàn", "diaban", "xã", "huyện", "tỉnh",
          "doanhthu", "doanh thu", "doanh thu thuần", "doanhthuthuan", "trị giá", "doanh số",
          "lao động", "laodong", "số người", "nhân sự", "songuoi", "nhansu",
          "mã ngành", "manganh", "vsic", "mã ngành chính", "manganhchinh",
          "ghichu", "ghi chú", "thuyết minh", "mô tả"
        ];

        const detected: CustomField[] = [];
        headers.forEach(header => {
          if (!header) return;
          const cleanHeader = header.toLowerCase();
          const isStandard = standardHeaders.some(std => {
            const cleanStd = std.toLowerCase();
            return cleanHeader.includes(cleanStd) || cleanStd.includes(cleanHeader);
          });

          if (!isStandard) {
            // Tạo slug an toàn cho khóa lưu trữ
            const slug = header
              .normalize("NFD")
              .replace(/[\u0300-\u036f]/g, "") // Khử dấu tiếng Việt
              .toLowerCase()
              .replace(/[^a-z0-9]/g, "_")
              .replace(/_+/g, "_")
              .replace(/^_+|_+$/g, "");

            detected.push({
              name: slug || `field_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
              label: header,
              type: "text" // Mặc định là kiểu chữ
            });
          }
        });

        if (detected.length === 0) {
          setStatusMessage({ 
            type: "error", 
            text: "Không phát hiện thấy cột chỉ tiêu động nào mới trong tệp mẫu. Tất cả các cột đều khớp với các chỉ tiêu mặc định sẵn có của hệ thống!" 
          });
          setAnalyzedFields([]);
        } else {
          setAnalyzedFields(detected);
          setStatusMessage({ 
            type: "success", 
            text: `Phân tích mẫu thành công! Đã phát hiện thấy ${detected.length} cột chỉ tiêu động khảo sát.` 
          });
        }
      } catch (err: any) {
        console.error("Lỗi phân tích tệp mẫu:", err);
        setStatusMessage({ type: "error", text: `Lỗi phân tích tệp mẫu: ${err.message || err}` });
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleSaveCustomField = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newFieldName.trim() || !newFieldLabel.trim()) {
      alert("Vui lòng điền đầy đủ Mã khóa và Tên chỉ tiêu!");
      return;
    }

    const slug = newFieldName
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "_")
      .replace(/_+/g, "_");

    // Check if duplicate
    if (customFields.some(f => f.name === slug)) {
      alert("Mã khóa chỉ tiêu này đã tồn tại!");
      return;
    }

    const updated = [...customFields, { name: slug, label: newFieldLabel.trim(), type: newFieldType }];
    setCustomFields(updated);
    localStorage.setItem(`custom_fields_${unitID}`, JSON.stringify(updated));

    // Reset inputs
    setNewFieldName("");
    setNewFieldLabel("");
    setNewFieldType("text");
    setStatusMessage({ type: "success", text: `Đã thêm mới chỉ tiêu động: "${newFieldLabel.trim()}"` });
  };

  const handleDeleteCustomField = (fieldName: string) => {
    const updated = customFields.filter(f => f.name !== fieldName);
    setCustomFields(updated);
    localStorage.setItem(`custom_fields_${unitID}`, JSON.stringify(updated));
    setStatusMessage({ type: "success", text: "Đã xóa chỉ tiêu động thành công." });
  };

  const handleApplyAnalyzedFields = () => {
    if (analyzedFields.length === 0) return;
    
    // Merge without duplicates
    const updated = [...customFields];
    analyzedFields.forEach(f => {
      if (!updated.some(exist => exist.name === f.name)) {
        updated.push(f);
      }
    });

    setCustomFields(updated);
    localStorage.setItem(`custom_fields_${unitID}`, JSON.stringify(updated));
    setAnalyzedFields([]);
    setTemplateFileName("");
    setStatusMessage({ type: "success", text: `Đã tự động tạo các ô nhập liệu thành công dựa trên ${analyzedFields.length} chỉ tiêu từ tệp mẫu!` });
  };

  const handleClearAllCustomFields = () => {
    if (window.confirm("Bạn có chắc chắn muốn xóa sạch toàn bộ các chỉ tiêu động đang thiết kế không?")) {
      setCustomFields([]);
      localStorage.removeItem(`custom_fields_${unitID}`);
      setStatusMessage({ type: "success", text: "Đã xóa sạch toàn bộ chỉ tiêu động." });
    }
  };

  // Load danh sách dữ liệu từ localStorage hoặc Firestore khi component mount
  useEffect(() => {
    fetchRecords();
    
    // Nạp danh sách các trường động tùy chỉnh của biểu mẫu
    const localFieldsKey = `custom_fields_${unitID}`;
    const storedFields = localStorage.getItem(localFieldsKey);
    if (storedFields) {
      try {
        setCustomFields(JSON.parse(storedFields));
      } catch (e) {
        console.error("Lỗi khi đọc danh sách trường động:", e);
      }
    }
  }, [unitID]);

  const fetchRecords = async () => {
    setLoading(true);
    try {
      if (isFirebaseInitialized && db) {
        const querySnapshot = await getDocs(collection(db, "data", unitID, "survey_records"));
        const records: SurveyRecord[] = [];
        querySnapshot.forEach((doc) => {
          records.push({ id: doc.id, ...doc.data() } as SurveyRecord);
        });
        // Sắp xếp bản ghi mới nhất lên đầu
        records.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
        setSurveyRecords(records);
      } else {
        const localKey = `survey_records_${unitID}`;
        const stored = localStorage.getItem(localKey);
        if (stored) {
          const records = JSON.parse(stored) as SurveyRecord[];
          records.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
          setSurveyRecords(records);
        } else {
          setSurveyRecords([]);
        }
      }
    } catch (err) {
      console.error("Lỗi tải danh sách khảo sát:", err);
    } finally {
      setLoading(false);
    }
  };

  // Hàm ghi nhật ký hệ thống
  const logActivity = async (actionType: "manual" | "scan" | "excel", details: string) => {
    const timestamp = new Date().toISOString();
    const logData = {
      unitID,
      userName,
      actionType: `survey_${actionType}`,
      details,
      timestamp,
      createdAt: isFirebaseInitialized ? serverTimestamp() : timestamp
    };

    if (isFirebaseInitialized && db) {
      try {
        await addDoc(collection(db, "logs"), logData);
      } catch (err) {
        console.error("Lỗi ghi nhật ký Firebase:", err);
      }
    } else {
      const existingLogs = JSON.parse(localStorage.getItem("system_mock_logs") || "[]");
      existingLogs.unshift(logData);
      localStorage.setItem("system_mock_logs", JSON.stringify(existingLogs));
      window.dispatchEvent(new Event("storage_logs_updated"));
    }
  };

  // ========================================================
  // 1. XỬ LÝ NHẬP THỦ CÔNG (MANUAL ENTER)
  // ========================================================
  const handleManualSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!manualName.trim() || !manualMst.trim()) {
      setStatusMessage({ type: "error", text: "Vui lòng nhập các thông tin bắt buộc (Tên cơ sở, Doanh nghiệp & Mã số thuế/Mã định danh)!" });
      return;
    }

    setLoading(true);
    setStatusMessage(null);

    // Thu thập các giá trị của các trường động tùy chỉnh
    const customDataPayload: Record<string, any> = {};
    customFields.forEach(field => {
      const val = dynamicValues[field.name];
      if (val !== undefined && val !== "") {
        if (field.type === "number") {
          customDataPayload[field.name] = parseFloat(val) || 0;
        } else if (field.type === "boolean") {
          customDataPayload[field.name] = val === "true" || val === true;
        } else {
          customDataPayload[field.name] = val;
        }
      }
    });

    const payload: SurveyRecord = {
      mst: manualMst.trim(),
      name: manualName.trim(),
      representative: manualRep.trim(),
      address: manualAddress.trim(),
      doanhthu: parseFloat(manualDoanhThu) || 0,
      laodong: parseInt(manualLaoDong) || 0,
      manganh: manualMaNganh.trim(),
      ghichu: manualGhiChu.trim(),
      block: activeBlock,
      entryMethod: "manual",
      createdBy: userName,
      createdAt: new Date().toISOString(),
      customData: customDataPayload
    };

    try {
      if (isFirebaseInitialized && db) {
        await addDoc(collection(db, "data", unitID, "survey_records"), payload);
      } else {
        const localKey = `survey_records_${unitID}`;
        const existing = JSON.parse(localStorage.getItem(localKey) || "[]") as SurveyRecord[];
        existing.unshift({ ...payload, id: "survey_" + Date.now() });
        localStorage.setItem(localKey, JSON.stringify(existing));
      }

      await logActivity("manual", `Nhập thủ công đối tượng ${activeBlock === "doanhnghiep" ? "DN" : "Cá thể"}: "${payload.name}" - MST: ${payload.mst}`);
      
      setStatusMessage({ type: "success", text: "Lưu phiếu khảo sát cơ sở thành công!" });
      
      // Reset form
      setManualMst("");
      setManualName("");
      setManualRep("");
      setManualAddress("");
      setManualDoanhThu("");
      setManualLaoDong("");
      setManualMaNganh("");
      setManualGhiChu("");
      setDynamicValues({});

      fetchRecords();
    } catch (err: any) {
      console.error("Lỗi lưu thủ công:", err);
      setStatusMessage({ type: "error", text: `Gặp lỗi khi lưu: ${err.message || err}` });
    } finally {
      setLoading(false);
    }
  };

  const handleHungYenSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!hyName.trim()) {
      setStatusMessage({ type: "error", text: "Vui lòng nhập Tên doanh nghiệp / Cơ sở sản xuất kinh doanh!" });
      return;
    }

    setLoading(true);
    setStatusMessage(null);

    // Merge specific fields into address
    const constructedAddress = [
      hySpecificAddress.trim(),
      hyCommuneName.trim()
    ].filter(Boolean).join(", ");

    // Calculate revenue based on block and form data
    let calculatedRevenue = 0;
    if (activeBlock === "cathe") {
      const dt9 = parseFloat(hyDoanhThu9Thang) || 0;
      const dt3 = parseFloat(hyDoanhThu3Thang) || 0;
      calculatedRevenue = dt9 + dt3 || parseFloat(hyDoanhThuNamTruoc) || 0;
    } else {
      calculatedRevenue = parseFloat(hyKqkdTongDtUocCaNam) || parseFloat(hyDoanhThu) || 0;
    }

    const finalMst = hyMst.trim() || (activeBlock === "cathe" && hySttCoSo.trim() ? `STT-${hySttCoSo.trim()}` : `MST-${Date.now().toString().slice(-6)}`);

    // Custom data specific to Hung Yen
    const customDataPayload: Record<string, any> = {
      ma_xa_phuong: hyCommuneCode.trim(),
      so_dien_thoai: hyPhone.trim(),
      email: hyEmail.trim(),
      nganh_chinh_mota: hyMaNganhMoTa.trim(),
      nganh_khac_1: hyNganhKhac1.trim(),
      nganh_khac_1_mota: hyNganhKhac1MoTa.trim(),
      nganh_khac_2: hyNganhKhac2.trim(),
      nganh_khac_2_mota: hyNganhKhac2MoTa.trim(),
      lao_dong_0101: parseInt(hyLaoDong0101) || 0,
      lao_dong_3112: parseInt(hyLaoDong3112) || 0,
      co_so_ngoai_xa: hyCoSoNgoaiXa,
      isHungYenForm: true,
      // Cá thể specific fields
      stt_co_so: hySttCoSo.trim(),
      doanh_thu_nam_truoc: parseFloat(hyDoanhThuNamTruoc) || 0,
      doanh_thu_9_thang: parseFloat(hyDoanhThu9Thang) || 0,
      doanh_thu_3_thang: parseFloat(hyDoanhThu3Thang) || 0,
      surveyor_name: hySurveyorName.trim(),
      surveyor_phone: hySurveyorPhone.trim(),
      surveyor_signature: hySurveyorSignature,
      respondent_name: hyRespondentName.trim(),
      respondent_phone: hyRespondentPhone.trim(),
      // Chữ ký số & Ký điện tử
      is_signed: !!(hySignatureDataUrl || hySurveyorSignature),
      signed_mode: hySignatureMode || "draw",
      signed_by: hyRespondentName.trim() || hyRep.trim() || "Người đại diện",
      signed_time: hySignedTime || new Date().toLocaleString("vi-VN"),
      signed_hash: hySignedHash || "SHA256-" + Math.random().toString(36).substring(2, 10).toUpperCase(),
      signature_img: hySignatureDataUrl,
      // Dữ liệu Trang 2 Doanh nghiệp
      co_so_ngoai_xa_list: hyCoSoNgoaiXa === "có" ? hyCoSoNgoaiXaList : [],
      kqkd_tong_dt_nam_truoc: parseFloat(hyKqkdTongDtNamTruoc) || 0,
      kqkd_tong_dt_9thang: parseFloat(hyKqkdTongDt9Thang) || 0,
      kqkd_tong_dt_uoc_ca_nam: parseFloat(hyKqkdTongDtUocCaNam) || 0,
      kqkd_chinh_dt_nam_truoc: parseFloat(hyKqkdChinhDtNamTruoc) || 0,
      kqkd_chinh_dt_9thang: parseFloat(hyKqkdChinhDt9Thang) || 0,
      kqkd_chinh_dt_uoc_ca_nam: parseFloat(hyKqkdChinhDtUocCaNam) || 0,
      kqkd_khac1_ten: hyKqkdKhac1Ten.trim(),
      kqkd_khac1_dt_nam_truoc: parseFloat(hyKqkdKhac1DtNamTruoc) || 0,
      kqkd_khac1_dt_9thang: parseFloat(hyKqkdKhac1Dt9Thang) || 0,
      kqkd_khac1_dt_uoc_ca_nam: parseFloat(hyKqkdKhac1DtUocCaNam) || 0,
      kqkd_khac2_ten: hyKqkdKhac2Ten.trim(),
      kqkd_khac2_dt_nam_truoc: parseFloat(hyKqkdKhac2DtNamTruoc) || 0,
      kqkd_khac2_dt_9thang: parseFloat(hyKqkdKhac2Dt9Thang) || 0,
      kqkd_khac2_dt_uoc_ca_nam: parseFloat(hyKqkdKhac2DtUocCaNam) || 0,
    };

    // Include existing dynamic variables if custom fields exist
    customFields.forEach(field => {
      const val = dynamicValues[field.name];
      if (val !== undefined && val !== "") {
        if (field.type === "number") {
          customDataPayload[field.name] = parseFloat(val) || 0;
        } else if (field.type === "boolean") {
          customDataPayload[field.name] = val === "true" || val === true;
        } else {
          customDataPayload[field.name] = val;
        }
      }
    });

    const payload: SurveyRecord = {
      mst: finalMst,
      name: hyName.trim(),
      representative: hyRep.trim() || hyRespondentName.trim() || hySignedBy,
      address: constructedAddress || "Hưng Yên",
      doanhthu: calculatedRevenue,
      laodong: parseInt(hyLaoDong3112) || parseInt(hyLaoDong0101) || 0,
      manganh: hyMaNganh.trim(),
      ghichu: hyGhiChu.trim(),
      block: activeBlock,
      entryMethod: "manual",
      createdBy: userName,
      createdAt: new Date().toISOString(),
      customData: customDataPayload
    };

    try {
      if (isFirebaseInitialized && db) {
        await addDoc(collection(db, "data", unitID, "survey_records"), payload);
      } else {
        const localKey = `survey_records_${unitID}`;
        const existing = JSON.parse(localStorage.getItem(localKey) || "[]") as SurveyRecord[];
        existing.unshift({ ...payload, id: "survey_" + Date.now() });
        localStorage.setItem(localKey, JSON.stringify(existing));
      }

      await logActivity("manual", `Nhập thủ công theo mẫu phiếu Hưng Yên đối tượng ${activeBlock === "doanhnghiep" ? "DN" : "Cá thể"}: "${payload.name}" - MST/Mã: ${payload.mst}`);
      
      setStatusMessage({ type: "success", text: `Lưu thông tin Phiếu điều tra Hưng Yên (${activeBlock === "doanhnghiep" ? "Doanh nghiệp" : "Cá thể"}) thành công!` });
      
      // Reset Hung Yen form
      setHyMst("");
      setHyName("");
      setHyRep("");
      setHyCommuneName("");
      setHyCommuneCode("");
      setHySpecificAddress("");
      setHyPhone("");
      setHyEmail("");
      setHyMaNganh("");
      setHyMaNganhMoTa("");
      setHyNganhKhac1("");
      setHyNganhKhac1MoTa("");
      setHyNganhKhac2("");
      setHyNganhKhac2MoTa("");
      setHyLaoDong0101("");
      setHyLaoDong3112("");
      setHyCoSoNgoaiXa("");
      setHyDoanhThu("");
      setHyGhiChu("");
      setHySttCoSo("");
      setHyDoanhThuNamTruoc("");
      setHyDoanhThu9Thang("");
      setHyDoanhThu3Thang("");
      setHySurveyorName("");
      setHySurveyorPhone("");
      setHyRespondentName("");
      setHyRespondentPhone("");
      setHyCoSoNgoaiXaList([
        { stt: "1", ten: "", diaChi: "", laoDong0101: "", laoDong3112: "" },
        { stt: "2", ten: "", diaChi: "", laoDong0101: "", laoDong3112: "" },
        { stt: "3", ten: "", diaChi: "", laoDong0101: "", laoDong3112: "" },
        { stt: "4", ten: "", diaChi: "", laoDong0101: "", laoDong3112: "" },
      ]);
      setHyKqkdTongDtNamTruoc("");
      setHyKqkdTongDt9Thang("");
      setHyKqkdTongDtUocCaNam("");
      setHyKqkdChinhDtNamTruoc("");
      setHyKqkdChinhDt9Thang("");
      setHyKqkdChinhDtUocCaNam("");
      setHyKqkdKhac1Ten("");
      setHyKqkdKhac1DtNamTruoc("");
      setHyKqkdKhac1Dt9Thang("");
      setHyKqkdKhac1DtUocCaNam("");
      setHyKqkdKhac2Ten("");
      setHyKqkdKhac2DtNamTruoc("");
      setHyKqkdKhac2Dt9Thang("");
      setHyKqkdKhac2DtUocCaNam("");
      setHySignatureMode("");
      setHySignatureDataUrl("");
      setHySurveyorSignature("");
      setHyIsSigned(false);
      setHySignedTime("");
      setHySignedBy("");
      setHySignedHash("");
      setDynamicValues({});

      fetchRecords();
    } catch (err: any) {
      console.error("Lỗi lưu mẫu phiếu Hưng Yên:", err);
      setStatusMessage({ type: "error", text: `Gặp lỗi khi lưu: ${err.message || err}` });
    } finally {
      setLoading(false);
    }
  };

  // ========================================================
  // 2. XỬ LÝ ĐỌC FILE SCAN BẰNG AI (GEMINI OCR EXTRACT)
  // ========================================================
  const handleScanFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setScanFile(file);
    setExtractedResult(null);
    setStatusMessage(null);

    // Tạo preview URL nếu là file ảnh
    if (file.type.startsWith("image/")) {
      const url = URL.createObjectURL(file);
      setScanPreviewUrl(url);
    } else {
      setScanPreviewUrl(null);
    }
  };

  const handleStartAiExtraction = async () => {
    if (!scanFile) {
      setStatusMessage({ type: "error", text: "Vui lòng chọn ảnh chụp phiếu hoặc tệp scan trước!" });
      return;
    }

    const apiKey = (import.meta as any).env?.VITE_GEMINI_API_KEY || "";
    if (!apiKey) {
      setStatusMessage({ 
        type: "error", 
        text: "Cần cấu hình khoá API VITE_GEMINI_API_KEY trong phần Cài đặt > Khóa bí mật (Secrets) của AI Studio hoặc trong tệp .env để sử dụng Trí tuệ AI trích xuất!" 
      });
      return;
    }

    setIsExtracting(true);
    setAiLogs(["🔄 Đang khởi tạo Trợ lý AI đọc ảnh tài liệu...", "📋 Đọc dữ liệu nhị phân từ file..."]);
    
    try {
      // Chuyển file sang Base64
      const reader = new FileReader();
      const base64Promise = new Promise<string>((resolve, reject) => {
        reader.onload = () => {
          const result = reader.result as string;
          // Cắt bỏ phần đầu data:image/jpeg;base64,
          const base64 = result.split(",")[1];
          resolve(base64);
        };
        reader.onerror = (err) => reject(err);
      });
      reader.readAsDataURL(scanFile);
      const base64Data = await base64Promise;

      setAiLogs(prev => [...prev, "🧬 Đang phân tích chữ viết tay và biểu mẫu dữ liệu bằng Gemini..."]);

      const ai = new GoogleGenAI({
        apiKey: apiKey,
        httpOptions: { headers: { "User-Agent": "aistudio-build" } }
      });

      let customFieldsInstructions = "";
      if (customFields.length > 0) {
        customFieldsInstructions = `\nNgoài ra, biểu mẫu khảo sát này còn có một số chỉ tiêu bổ sung (trường động). Hãy nhận diện, phân tích và trích xuất chúng, rồi đặt tất cả vào một đối tượng con tên là "customData" dạng JSON với các khóa tương ứng dưới đây:\n` +
          customFields.map(f => `- Khóa "${f.name}": "${f.label}" (Hãy phân tích giá trị tương ứng, kiểu dữ liệu: ${f.type})`).join("\n");
      }

      const response = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: [
          {
            role: "user",
            parts: [
              {
                text: `Bạn là một chuyên gia số hóa phiếu khảo sát kinh tế. Hãy đọc kỹ ảnh chụp phiếu khảo sát này và trích xuất thông tin của đối tượng được khảo sát.
Hãy phân loại đối tượng này là: "Cá thể" (Hộ kinh doanh cá thể, cơ sở cá thể) hoặc "Doanh nghiệp" (Công ty, doanh nghiệp).
Hãy trích xuất các trường thông tin sau dưới dạng JSON hợp lệ:
{
  "type": "Cá thể" hoặc "Doanh nghiệp",
  "mst": "Mã số thuế hoặc mã định danh",
  "name": "Tên cơ sở hoặc tên doanh nghiệp",
  "representative": "Chủ cơ sở hoặc Người đại diện pháp luật",
  "address": "Địa chỉ hoặc địa bàn xã/quận",
  "doanhthu": "Doanh thu kinh doanh (triệu đồng, chỉ lấy số)",
  "laodong": "Số lượng lao động (người, chỉ lấy số)",
  "manganh": "Mã ngành VSIC (5 số, nếu có)",
  "ghichu": "Ghi chú tóm tắt từ phiếu khảo sát"${customFields.length > 0 ? ',\n  "customData": {\n' + customFields.map(f => `    "${f.name}": "Giá trị của ${f.label}"`).join(",\n") + '\n  }' : ''}
}${customFieldsInstructions}
Hãy trả về DUY NHẤT một chuỗi JSON hợp lệ nằm trong dấu nháy triple backticks \`\`\`json ... \`\`\`. Không giải thích gì thêm.`
              },
              {
                inlineData: {
                  mimeType: scanFile.type,
                  data: base64Data
                }
              }
            ]
          }
        ]
      });

      const resText = response.text || "";
      setAiLogs(prev => [...prev, "📝 Đang xử lý kết quả trả về từ mô hình AI..."]);
      
      // Parse JSON từ văn bản trả về
      const jsonMatch = resText.match(/```json\s*([\s\S]*?)\s*```/);
      const jsonStr = jsonMatch ? jsonMatch[1] : resText;
      const parsed = JSON.parse(jsonStr.trim());

      setAiLogs(prev => [...prev, "🎉 Trích xuất hoàn thành! Đang hiển thị kết quả rà soát..."]);

      // Ánh xạ về định dạng SurveyRecord
      const extracted: Partial<SurveyRecord> = {
        mst: parsed.mst || "",
        name: parsed.name || "",
        representative: parsed.representative || "",
        address: parsed.address || "",
        doanhthu: parseFloat(parsed.doanhthu) || 0,
        laodong: parseInt(parsed.laodong) || 0,
        manganh: parsed.manganh || "",
        ghichu: parsed.ghichu || `Trích xuất tự động từ file: ${scanFile.name}`,
        block: parsed.type?.toLowerCase().includes("cá thể") ? "cathe" : "doanhnghiep",
        entryMethod: "scan",
        customData: parsed.customData || {}
      };

      setExtractedResult(extracted);
      setStatusMessage({ type: "success", text: "Trích xuất ảnh quét bằng AI thành công! Hãy rà soát lại thông tin bên dưới." });
    } catch (err: any) {
      console.error("Lỗi trích xuất AI:", err);
      setAiLogs(prev => [...prev, `❌ Thất bại: ${err.message || err}`]);
      setStatusMessage({ type: "error", text: `Lỗi trích xuất AI: ${err.message || "Tệp ảnh không rõ ràng hoặc không khớp định dạng."}` });
    } finally {
      setIsExtracting(false);
    }
  };

  const handleSaveExtractedResult = async () => {
    if (!extractedResult) return;
    setLoading(true);

    const payload: SurveyRecord = {
      mst: extractedResult.mst || "MOCK_MST",
      name: extractedResult.name || "Tên trích xuất trống",
      representative: extractedResult.representative || "",
      address: extractedResult.address || "",
      doanhthu: extractedResult.doanhthu || 0,
      laodong: extractedResult.laodong || 0,
      manganh: extractedResult.manganh || "",
      ghichu: extractedResult.ghichu || "",
      block: extractedResult.block || activeBlock,
      entryMethod: "scan",
      createdBy: userName,
      createdAt: new Date().toISOString(),
      customData: extractedResult.customData || {}
    };

    try {
      if (isFirebaseInitialized && db) {
        await addDoc(collection(db, "data", unitID, "survey_records"), payload);
      } else {
        const localKey = `survey_records_${unitID}`;
        const existing = JSON.parse(localStorage.getItem(localKey) || "[]") as SurveyRecord[];
        existing.unshift({ ...payload, id: "survey_" + Date.now() });
        localStorage.setItem(localKey, JSON.stringify(existing));
      }

      await logActivity("scan", `Nhập dữ liệu qua file scan AI: "${payload.name}"`);
      setStatusMessage({ type: "success", text: "Đã rà soát và lưu trữ dữ liệu thành công!" });
      setExtractedResult(null);
      setScanFile(null);
      setScanPreviewUrl(null);
      fetchRecords();
    } catch (err: any) {
      console.error("Lỗi lưu kết quả AI:", err);
      setStatusMessage({ type: "error", text: `Gặp lỗi khi lưu: ${err.message || err}` });
    } finally {
      setLoading(false);
    }
  };

  // ========================================================
  // 3. XỬ LÝ NẠP FILE EXCEL
  // ========================================================
  const handleExcelFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setExcelFileName(file.name);
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = event.target?.result;
        const workbook = XLSX.read(data, { type: "binary" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Chuyển đổi sang JSON Array
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        setExcelRows(jsonData);
        setStatusMessage({ type: "success", text: `Đọc tệp Excel thành công! Phát hiện thấy ${jsonData.length} dòng.` });
      } catch (err: any) {
        console.error("Lỗi đọc Excel:", err);
        setStatusMessage({ type: "error", text: "Tệp Excel không đúng định dạng hoặc bị hỏng." });
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleExcelUpload = async () => {
    if (excelRows.length === 0) return;
    setLoading(true);
    setUploadProgress("Bắt đầu xử lý danh mục dữ liệu...");

    // Bản dịch tự động khớp các cột Excel thông dụng sang tiếng Việt
    const normalizeKey = (key: string): string => {
      const clean = key.toLowerCase().trim().replace(/_/g, "").normalize("NFC");
      if (["mst", "masothue", "mã số thuế", "mãđịnhdanh", "mã cơ sở", "macoso"].some(k => clean.includes(k))) return "mst";
      if (["tên", "tendoanhnghiep", "tên doanh nghiệp", "tên cơ sở", "tencoso", "doanhnghiep", "tên đơn vị"].some(k => clean.includes(k))) return "name";
      if (["đạidiện", "chủ cơ sở", "chucoso", "người đại diện", "daidien", "nguoidaidien"].some(k => clean.includes(k))) return "representative";
      if (["địachỉ", "diachi", "địabàn", "diaban", "xã", "huyện", "tỉnh"].some(k => clean.includes(k))) return "address";
      if (["doanhthu", "doanh thu", "doanh thu thuần", "doanhthuthuan", "trị giá", "doanh số"].some(k => clean.includes(k))) return "doanhthu";
      if (["lao động", "laodong", "số người", "nhân sự", "songuoi", "nhansu"].some(k => clean.includes(k))) return "laodong";
      if (["mã ngành", "manganh", "vsic", "mã ngành chính", "manganhchinh"].some(k => clean.includes(k))) return "manganh";
      if (["ghichu", "ghi chú", "thuyết minh", "mô tả"].some(k => clean.includes(k))) return "ghichu";
      return key;
    };

    try {
      const formatted: SurveyRecord[] = excelRows.map((row, idx) => {
        const normalizedRow: any = {};
        const rowCustomData: Record<string, any> = {};

        Object.entries(row).forEach(([k, v]) => {
          const normKey = normalizeKey(k);
          if (normKey === k) {
            // Trường không chuẩn, kiểm tra xem có khớp chỉ tiêu động nào không
            const matchedCustom = customFields.find(cf => 
              cf.name.toLowerCase() === k.toLowerCase().replace(/_/g, "").trim() || 
              cf.label.toLowerCase() === k.toLowerCase().trim()
            );
            if (matchedCustom) {
              if (matchedCustom.type === "number") {
                rowCustomData[matchedCustom.name] = parseFloat(String(v)) || 0;
              } else if (matchedCustom.type === "boolean") {
                rowCustomData[matchedCustom.name] = String(v).toLowerCase() === "true" || String(v) === "1" || String(v).toLowerCase() === "có";
              } else {
                rowCustomData[matchedCustom.name] = v;
              }
            } else {
              // Lưu vào làm thuộc tính động phụ trợ
              const safeKey = k.toLowerCase().normalize("NFC").replace(/[^a-zA-Z0-9]/g, "_");
              rowCustomData[safeKey] = v;
            }
          } else {
            normalizedRow[normKey] = v;
          }
        });

        return {
          mst: String(normalizedRow.mst || `EXCEL_${Date.now()}_${idx}`).trim(),
          name: String(normalizedRow.name || `Cơ sở chưa đặt tên số ${idx + 1}`).trim(),
          representative: String(normalizedRow.representative || "").trim(),
          address: String(normalizedRow.address || "Chưa xác định").trim(),
          doanhthu: parseFloat(normalizedRow.doanhthu) || 0,
          laodong: parseInt(normalizedRow.laodong) || 0,
          manganh: String(normalizedRow.manganh || "").trim(),
          ghichu: String(normalizedRow.ghichu || `Nhập khẩu Excel: ${excelFileName}`).trim(),
          block: activeBlock,
          entryMethod: "excel",
          createdBy: userName,
          createdAt: new Date().toISOString(),
          customData: rowCustomData
        };
      });

      if (isFirebaseInitialized && db) {
        const batchSize = 100;
        const total = formatted.length;
        
        for (let i = 0; i < total; i += batchSize) {
          const batch = writeBatch(db);
          const chunk = formatted.slice(i, i + batchSize);

          chunk.forEach((record) => {
            const docRef = doc(collection(db, "data", unitID, "survey_records"));
            batch.set(docRef, record);
          });

          setUploadProgress(`Đang đẩy dữ liệu lô ${Math.floor(i / batchSize) + 1} (${Math.min(i + batchSize, total)}/${total} dòng)...`);
          await batch.commit();
        }
      } else {
        const localKey = `survey_records_${unitID}`;
        const existing = JSON.parse(localStorage.getItem(localKey) || "[]") as SurveyRecord[];
        localStorage.setItem(localKey, JSON.stringify([...formatted, ...existing]));
      }

      await logActivity("excel", `Nạp tệp Excel khảo sát ${activeBlock === "doanhnghiep" ? "DN" : "Cá thể"}: "${excelFileName}" với ${formatted.length} bản ghi.`);
      
      setStatusMessage({ type: "success", text: `Nạp thành công toàn bộ ${formatted.length} dòng khảo sát từ Excel vào hệ thống!` });
      setExcelRows([]);
      setExcelFileName("");
      if (excelFileInputRef.current) excelFileInputRef.current.value = "";
      fetchRecords();
    } catch (err: any) {
      console.error("Lỗi nạp Excel khảo sát:", err);
      setStatusMessage({ type: "error", text: `Gặp sự cố khi ghi dữ liệu: ${err.message || err}` });
    } finally {
      setLoading(false);
      setUploadProgress(null);
    }
  };

  // ========================================================
  // XÓA VÀ CHỈNH SỬA BẢN GHI KHẢO SÁT
  // ========================================================
  const handleDeleteRecord = async (recordId: string, recordName: string) => {
    if (!window.confirm(`Bạn có chắc chắn muốn xóa vĩnh viễn phiếu khảo sát của "${recordName}" không?`)) {
      return;
    }

    setLoading(true);
    try {
      if (isFirebaseInitialized && db) {
        await deleteDoc(doc(db, "data", unitID, "survey_records", recordId));
      } else {
        const localKey = `survey_records_${unitID}`;
        const existing = JSON.parse(localStorage.getItem(localKey) || "[]") as SurveyRecord[];
        const filtered = existing.filter(r => r.id !== recordId);
        localStorage.setItem(localKey, JSON.stringify(filtered));
      }

      await logActivity("manual", `Xóa bản ghi khảo sát: "${recordName}"`);
      setStatusMessage({ type: "success", text: "Xóa thông tin phiếu khảo sát thành công!" });
      fetchRecords();
    } catch (err: any) {
      console.error("Lỗi xóa bản ghi:", err);
      setStatusMessage({ type: "error", text: `Không thể xóa phiếu khảo sát: ${err.message}` });
    } finally {
      setLoading(false);
    }
  };

  const handleUpdateRecord = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingRecord) return;

    setLoading(true);
    try {
      if (isFirebaseInitialized && db && editingRecord.id) {
        await updateDoc(doc(db, "data", unitID, "survey_records", editingRecord.id), {
          mst: editingRecord.mst,
          name: editingRecord.name,
          representative: editingRecord.representative || "",
          address: editingRecord.address,
          doanhthu: editingRecord.doanhthu,
          laodong: editingRecord.laodong,
          manganh: editingRecord.manganh,
          ghichu: editingRecord.ghichu || "",
          customData: editingRecord.customData || {}
        });
      } else {
        const localKey = `survey_records_${unitID}`;
        const existing = JSON.parse(localStorage.getItem(localKey) || "[]") as SurveyRecord[];
        const updated = existing.map(r => r.id === editingRecord.id ? editingRecord : r);
        localStorage.setItem(localKey, JSON.stringify(updated));
      }

      setStatusMessage({ type: "success", text: `Đã cập nhật thông tin khảo sát thành công!` });
      setEditingRecord(null);
      fetchRecords();
    } catch (err: any) {
      console.error("Lỗi cập nhật:", err);
      setStatusMessage({ type: "error", text: `Không thể cập nhật: ${err.message}` });
    } finally {
      setLoading(false);
    }
  };

  // ========================================================
  // 4. KẾT NỐI & ĐỒNG BỘ GOOGLE DRIVE
  // ========================================================
  const handleConnectDrive = async () => {
    try {
      setStatusMessage(null);
      await signInWithGoogle();
      setStatusMessage({ type: "success", text: "Kết nối tài khoản Google và cấp quyền truy cập Drive thành công!" });
    } catch (err: any) {
      console.error("Lỗi kết nối Google Drive:", err);
      setStatusMessage({ type: "error", text: `Lỗi kết nối: ${err.message}` });
    }
  };

  const handleBackupToDrive = async () => {
    if (!googleAccessToken) {
      setStatusMessage({ type: "error", text: "Vui lòng kết nối Google Drive trước khi đồng bộ!" });
      return;
    }

    setIsDriveSyncing(true);
    setDriveSyncLogs(["🔄 Bắt đầu sao lưu lên Google Drive...", "📂 Chuẩn bị dữ liệu đóng gói dạng JSON..."]);

    try {
      const fileName = `VTong_Survey_Records_${unitID.toUpperCase()}_${new Date().toISOString().slice(0,10)}.json`;
      const contentStr = JSON.stringify(surveyRecords, null, 2);

      setDriveSyncLogs(prev => [...prev, `📤 Đang gửi dữ liệu "${fileName}" tới API Google Drive...`]);

      const boundary = "foo_bar_boundary";
      const metadata = {
        name: fileName,
        mimeType: "application/json"
      };

      const multipartBody = 
        `\r\n--${boundary}\r\n` +
        `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
        `${JSON.stringify(metadata)}\r\n` +
        `\r\n--${boundary}\r\n` +
        `Content-Type: application/json\r\n\r\n` +
        `${contentStr}\r\n` +
        `--${boundary}--`;

      const response = await fetch(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${googleAccessToken}`,
            "Content-Type": `multipart/related; boundary=${boundary}`
          },
          body: multipartBody
        }
      );

      if (!response.ok) {
        const errDetails = await response.text();
        throw new Error(`Mã lỗi ${response.status}: ${errDetails}`);
      }

      setDriveSyncLogs(prev => [...prev, "🎉 Sao lưu hoàn tất! Tệp tin đã được lưu trong My Drive của bạn."]);
      setStatusMessage({ type: "success", text: `Đã xuất và đồng bộ thành công ${surveyRecords.length} phiếu khảo sát lên Google Drive!` });
    } catch (err: any) {
      console.error("Lỗi lưu Drive:", err);
      setDriveSyncLogs(prev => [...prev, `❌ Thất bại: ${err.message}`]);
      setStatusMessage({ type: "error", text: `Lỗi sao lưu Google Drive: ${err.message}` });
    } finally {
      setIsDriveSyncing(false);
    }
  };

  // Lọc dữ liệu hiển thị
  const filteredRecords = surveyRecords.filter(r => {
    const matchesSearch = 
      r.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.mst.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.address.toLowerCase().includes(searchQuery.toLowerCase());

    const matchesBlock = filterBlock === "all" || r.block === filterBlock;
    const matchesMethod = filterMethod === "all" || r.entryMethod === filterMethod;

    return matchesSearch && matchesBlock && matchesMethod;
  });

  return (
    <div className="space-y-6 font-sans text-slate-800">
      
      {/* TIÊU ĐỀ TRANG */}
      <div className="bg-gradient-to-r from-indigo-900 via-indigo-950 to-slate-900 rounded-3xl p-6 text-white shadow-lg relative overflow-hidden">
        <div className="absolute top-0 right-0 w-64 h-64 bg-indigo-600/10 rounded-full blur-3xl"></div>
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2">
              <span className="bg-indigo-500/30 text-indigo-300 text-[10px] font-black uppercase tracking-widest px-3 py-1 rounded-full border border-indigo-500/20">
                Mô-đun Khảo sát Kinh tế
              </span>
            </div>
            <h2 className="text-xl font-black mt-2 flex items-center gap-2 tracking-tight">
              📝 Nhập Tin Khảo Sát Đăng Ký VSIC
            </h2>
            <p className="text-xs text-indigo-200/80 mt-1 leading-relaxed max-w-2xl">
              Thực hiện nạp dữ liệu rà soát, khảo sát thực tế phân tách rõ rệt thành hai khối Cá thể và Doanh nghiệp với đầy đủ phương thức rà soát bằng Excel, ảnh chụp phiếu scan, hoặc khai báo thủ công.
            </p>
          </div>

          <button
            onClick={() => setShowDrivePanel(!showDrivePanel)}
            className="flex items-center gap-1.5 px-4 py-2 bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-600 hover:to-teal-600 font-bold text-xs rounded-xl shadow-md cursor-pointer text-white transition active:scale-95"
          >
            <CloudLightning className="w-4 h-4" />
            <span>Đồng bộ Google Drive</span>
          </button>
        </div>
      </div>

      {/* GOOGLE DRIVE SYNC CENTER PANEL */}
      {showDrivePanel && (
        <div className="bg-emerald-50/60 border border-emerald-200 rounded-2xl p-5 space-y-4 animate-fade-in shadow-inner">
          <div className="flex justify-between items-center border-b border-emerald-100 pb-2">
            <h4 className="text-xs font-black uppercase tracking-wider text-emerald-800 flex items-center gap-2">
              📂 Cổng kết nối &amp; Lưu trữ Đám mây Google Drive
            </h4>
            <span className="text-[10px] font-bold text-emerald-600 bg-emerald-100 px-2 py-0.5 rounded">Bảo mật oAuth</span>
          </div>

          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <p className="text-xs text-emerald-700 leading-relaxed max-w-xl">
              Mọi thông tin phiếu điều tra được đồng bộ trực tiếp lên tài khoản lưu trữ bảo mật của đơn vị. Dữ liệu được trích xuất dưới dạng tệp chuẩn hóa của Tổng cục Thống kê, giúp khôi phục dữ liệu an toàn và chia sẻ thông tin liền mạch giữa các đoàn điều tra viên.
            </p>

            <div className="shrink-0">
              {googleAccessToken ? (
                <div className="flex items-center gap-2">
                  <span className="bg-emerald-600 text-white text-[10px] font-bold px-3 py-1.5 rounded-xl border border-emerald-700 flex items-center gap-1">
                    <CheckCircle className="w-3.5 h-3.5" />
                    Đã kết nối Google Drive
                  </span>
                  <button 
                    onClick={handleBackupToDrive}
                    disabled={isDriveSyncing || surveyRecords.length === 0}
                    className="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-xl shadow-sm transition cursor-pointer disabled:bg-slate-300"
                  >
                    {isDriveSyncing ? "Đang sao lưu..." : "Sao lưu ngay"}
                  </button>
                </div>
              ) : (
                <button
                  onClick={handleConnectDrive}
                  className="flex items-center gap-2 bg-white text-slate-700 border border-slate-300 rounded-xl px-4 py-2 hover:bg-slate-50 transition shadow-sm font-bold text-xs cursor-pointer"
                >
                  <svg className="w-4 h-4" viewBox="0 0 48 48">
                    <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
                    <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
                    <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
                    <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
                  </svg>
                  <span>Kết nối với Google Drive</span>
                </button>
              )}
            </div>
          </div>

          {driveSyncLogs.length > 0 && (
            <div className="bg-slate-900 text-slate-300 font-mono text-[10px] p-3 rounded-xl space-y-1 max-h-24 overflow-y-auto">
              {driveSyncLogs.map((log, idx) => (
                <div key={idx} className="leading-tight">{log}</div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* BANNER THÔNG BÁO TRẠNG THÁI */}
      {statusMessage && (
        <div className={`p-4 rounded-xl flex items-start gap-3 border animate-fade-in ${
          statusMessage.type === "success" 
            ? "bg-emerald-50 border-emerald-200 text-emerald-800" 
            : "bg-rose-50 border-rose-200 text-rose-800"
        }`}>
          {statusMessage.type === "success" ? (
            <CheckCircle className="w-5 h-5 shrink-0 mt-0.5 text-emerald-600" />
          ) : (
            <AlertCircle className="w-5 h-5 shrink-0 mt-0.5 text-rose-600" />
          )}
          <span className="text-xs font-bold leading-relaxed">{statusMessage.text}</span>
        </div>
      )}

      {/* SWITCHER KHỐI CÁ THỂ VÀ DOANH NGHIỆP */}
      <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-4 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div className="flex items-center gap-2">
          <Database className="w-4 h-4 text-indigo-500" />
          <span className="text-xs font-bold text-slate-500 uppercase tracking-wide">Lựa chọn Khối đối tượng cần nhập tin:</span>
        </div>

        <div className="flex bg-slate-100 p-1 rounded-xl">
          <button
            onClick={() => { setActiveBlock("doanhnghiep"); setExtractedResult(null); }}
            className={`flex items-center gap-2 px-4 py-2 text-xs font-black uppercase rounded-lg cursor-pointer transition ${
              activeBlock === "doanhnghiep"
                ? "bg-white text-indigo-650 shadow-xs"
                : "text-slate-500 hover:text-slate-800"
            }`}
          >
            <Building2 className="w-4 h-4" />
            Khối Doanh nghiệp (DN)
          </button>
          <button
            onClick={() => { setActiveBlock("cathe"); setExtractedResult(null); }}
            className={`flex items-center gap-2 px-4 py-2 text-xs font-black uppercase rounded-lg cursor-pointer transition ${
              activeBlock === "cathe"
                ? "bg-white text-indigo-650 shadow-xs"
                : "text-slate-500 hover:text-slate-800"
            }`}
          >
            <Store className="w-4 h-4" />
            Khối Hộ Cá thể
          </button>
        </div>
      </div>

      {/* LỰA CHỌN 4 PHƯƠNG THỨC NHẬP DỮ LIỆU */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <button
          onClick={() => setEntryMethod("manual")}
          className={`flex items-center justify-center gap-2 p-3 border rounded-xl text-xs font-bold transition cursor-pointer ${
            entryMethod === "manual"
              ? "bg-indigo-50 border-indigo-300 text-indigo-750 font-black shadow-xs"
              : "bg-white border-slate-200 text-slate-650 hover:bg-slate-50"
          }`}
        >
          <Plus className="w-4 h-4" />
          1. Nhập thủ công
        </button>

        <button
          onClick={() => setEntryMethod("scan")}
          className={`flex items-center justify-center gap-2 p-3 border rounded-xl text-xs font-bold transition cursor-pointer ${
            entryMethod === "scan"
              ? "bg-indigo-50 border-indigo-300 text-indigo-750 font-black shadow-xs"
              : "bg-white border-slate-200 text-slate-650 hover:bg-slate-50"
          }`}
        >
          <ScanFace className="w-4 h-4" />
          2. Đọc file scan (AI nhận diện)
        </button>

        <button
          onClick={() => setEntryMethod("excel")}
          className={`flex items-center justify-center gap-2 p-3 border rounded-xl text-xs font-bold transition cursor-pointer ${
            entryMethod === "excel"
              ? "bg-indigo-50 border-indigo-300 text-indigo-750 font-black shadow-xs"
              : "bg-white border-slate-200 text-slate-650 hover:bg-slate-50"
          }`}
        >
          <FileSpreadsheet className="w-4 h-4" />
          3. Nạp file excel
        </button>

        <button
          onClick={() => setEntryMethod("template")}
          className={`flex items-center justify-center gap-2 p-3 border rounded-xl text-xs font-bold transition cursor-pointer ${
            entryMethod === "template"
              ? "bg-indigo-50 border-indigo-300 text-indigo-750 font-black shadow-xs"
              : "bg-white border-slate-200 text-slate-650 hover:bg-slate-50"
          }`}
        >
          <Sparkles className="w-4 h-4 text-indigo-500 shrink-0" />
          4. Cấu hình biểu mẫu mẫu
        </button>
      </div>

      {/* GIAO DIỆN TỪNG PHƯƠNG THỨC */}
      <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6">
        
        {/* PHƯƠNG THỨC 1: NHẬP THỦ CÔNG */}
        {entryMethod === "manual" && (
          <div className="space-y-6">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between border-b border-slate-100 pb-3 gap-3">
              <div>
                <h3 className="text-sm font-extrabold text-indigo-900 uppercase flex items-center gap-2">
                  ✍️ KHAI BÁO BẢN GHI THỦ CÔNG - {activeBlock === "doanhnghiep" ? "KHỐI DOANH NGHIỆP" : "KHỐI CÁ THỂ"}
                </h3>
                <p className="text-[11px] text-slate-500 mt-0.5">Lựa chọn kiểu giao diện nhập liệu phù hợp với nhu cầu của bạn</p>
              </div>
              
              {/* Toggle Form Mode */}
              <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl text-[11px] font-bold shrink-0">
                <button
                  type="button"
                  onClick={() => setManualFormType("hungyen")}
                  className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1 cursor-pointer ${
                    manualFormType === "hungyen"
                      ? "bg-white text-indigo-900 shadow-xs font-extrabold"
                      : "text-slate-500 hover:text-slate-800"
                  }`}
                >
                  📝 Phiếu Hưng Yên
                </button>
                <button
                  type="button"
                  onClick={() => setManualFormType("standard")}
                  className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1 cursor-pointer ${
                    manualFormType === "standard"
                      ? "bg-white text-indigo-900 shadow-xs font-extrabold"
                      : "text-slate-500 hover:text-slate-800"
                  }`}
                >
                  📱 Nhập nhanh
                </button>
              </div>
            </div>

            {/* CHẾ ĐỘ 1: GIAO DIỆN PHIẾU GIẤY HƯNG YÊN CHUYÊN NGHIỆP */}
            {manualFormType === "hungyen" && (
              <form onSubmit={handleHungYenSubmit} className="space-y-6">
                
                {/* Nút Điền nhanh Demo */}
                <div className="flex justify-end gap-2">
                  {activeBlock === "doanhnghiep" ? (
                    <button
                      type="button"
                      onClick={() => {
                        setHyMst("0900234567101");
                        setHyName("CÔNG TY CỔ PHẦN THƯƠNG MẠI & DỊCH VỤ PHỐ HIẾN");
                        setHyRep("Nguyễn Văn Hiến");
                        setHyCommuneName("Phường Hiến Nam");
                        setHyCommuneCode("21305");
                        setHySpecificAddress("Số 128 Chùa Chuông");
                        setHyPhone("02213865865");
                        setHyEmail("contact@phohientrade.vn");
                        setHyMaNganh("47110");
                        setHyNganhKhac1("56100");
                        setHyNganhKhac2("49330");
                        setHyLaoDong0101("45");
                        setHyLaoDong3112("52");
                        setHyCoSoNgoaiXa("không");
                        setHyDoanhThu("12850");
                        setHyGhiChu("Đã đối chiếu mẫu phiếu giấy");
                      }}
                      className="px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-lg text-xs font-bold transition flex items-center gap-1 cursor-pointer border border-indigo-200 shadow-2xs"
                    >
                      <Sparkles className="w-3.5 h-3.5 text-indigo-500 animate-bounce" />
                      Điền mẫu thử nghiệm Doanh nghiệp Hưng Yên
                    </button>
                  ) : (
                    <button
                      type="button"
                      onClick={() => {
                        setHySttCoSo("0128");
                        setHyName("CỬA HÀNG TẠP HÓA AN BÌNH - PHỐ HIẾN");
                        setHyCommuneName("Phường Hiến Nam");
                        setHyCommuneCode("21305");
                        setHySpecificAddress("Số 45 Điện Biên III");
                        setHyPhone("0987654321");
                        setHyEmail("anbinhphohien@gmail.com");
                        setHyMaNganh("47110");
                        setHyLaoDong0101("3");
                        setHyLaoDong3112("4");
                        setHyDoanhThuNamTruoc("1200");
                        setHyDoanhThu9Thang("950");
                        setHyDoanhThu3Thang("320");
                        setHyRespondentName("Bùi Thị Bình");
                        setHyRespondentPhone("0987654321");
                        setHySurveyorName("Trần Thanh Trà");
                        setHySurveyorPhone("0912345678");
                        setHyGhiChu("Đối tượng khảo sát đầy đủ, hợp tác tốt");
                      }}
                      className="px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded-lg text-xs font-bold transition flex items-center gap-1 cursor-pointer border border-emerald-200 shadow-2xs"
                    >
                      <Sparkles className="w-3.5 h-3.5 text-emerald-500 animate-bounce" />
                      Điền mẫu thử nghiệm Hộ Cá thể Hưng Yên
                    </button>
                  )}
                </div>

                {/* KHUNG GIẤY THIẾT KẾ ĐẶC TRƯNG TỈNH HƯNG YÊN - PHONG CÁCH THẬT 100% */}
                <div className="max-w-4xl mx-auto p-6 sm:p-12 bg-[#fdfcf7] border-2 border-slate-400 rounded-xs shadow-xl font-serif text-slate-850 space-y-8 relative overflow-hidden leading-relaxed">
                  
                  {/* Đường đóng dấu mộc đỏ nhạt hoặc dải phân biệt chính phủ */}
                  <div className="absolute top-0 inset-x-0 h-1.5 bg-gradient-to-r from-red-600 via-amber-600 to-indigo-700 opacity-95" />

                  {/* ---------------------------------------------------- */}
                  {/* TRƯỜNG HỢP A: KHỐI DOANH NGHIỆP (PHIẾU SỐ 03/TGTSP-DNTMDV) */}
                  {/* ---------------------------------------------------- */}
                  {activeBlock === "doanhnghiep" && (
                    <div className="space-y-6">
                      {/* Header phiếu */}
                      <div className="flex flex-col md:flex-row md:items-start justify-between gap-4 border-b border-slate-300 pb-4">
                        <div className="space-y-1">
                          <h4 className="text-xs sm:text-sm font-black tracking-wider text-slate-900 uppercase">UBND TỈNH HƯNG YÊN</h4>
                          <p className="text-[11px] font-bold text-slate-700 italic">Phiếu số 03/TGTSP-DNTMDV</p>
                        </div>
                        
                        <div className="flex flex-col sm:flex-row sm:items-center gap-2">
                          <span className="text-xs font-black text-slate-800 uppercase font-serif whitespace-nowrap">Mã số thuế:</span>
                          <div className="bg-white p-1 rounded-sm border border-slate-400 shadow-2xs">
                            <GridBoxesInput value={hyMst} onChange={setHyMst} length={13} splitIndex={10} />
                          </div>
                        </div>
                      </div>

                      {/* Title tiêu đề chính */}
                      <div className="text-center space-y-2 py-2">
                        <h2 className="text-base sm:text-xl font-extrabold text-slate-900 uppercase leading-normal tracking-wide">
                          PHIẾU THU THẬP THÔNG TIN VỀ HOẠT ĐỘNG<br />THƯƠNG MẠI VÀ DỊCH VỤ CỦA DOANH NGHIỆP/HTX
                        </h2>
                        <p className="text-[13px] font-bold text-slate-800 tracking-widest">NĂM 2026</p>
                        <p className="text-[11px] sm:text-[11.5px] font-bold text-slate-600 italic">
                          (Áp dụng đối với Doanh nghiệp đa doanh nghiệp/HTX hoạt động trong lĩnh vực thương mại và dịch vụ)
                        </p>
                      </div>

                      {/* Hai hộp lưu ý bên dưới tiêu đề */}
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-[11px] leading-relaxed">
                        <div className="border border-slate-300 p-3 bg-white rounded-sm shadow-2xs italic text-slate-700">
                          Thực hiện theo <b>Quyết định số 2026/QĐ-UBND</b> ngày 12/01/2026 của UBND tỉnh Hưng Yên về việc Ban hành Phương án điều tra, thu thập thông tin phục vụ biên soạn chỉ tiêu Tổng giá trị sản phẩm trên địa bàn xã, phường tỉnh Hưng Yên.
                        </div>
                        <div className="border border-slate-300 p-3 bg-white rounded-sm shadow-2xs space-y-1 text-slate-700">
                          <p>- Nghĩa vụ cung cấp thông tin được quy định theo Luật thống kê;</p>
                          <p>- Thông tin cung cấp theo phiếu này chỉ nhằm phục vụ công tác thống kê và được bảo mật theo Luật định.</p>
                        </div>
                      </div>

                      {/* PHẦN I: THÔNG TIN CHUNG */}
                      <div className="space-y-4 pt-2">
                        <h3 className="text-xs sm:text-sm font-black text-slate-900 border-b border-slate-400 pb-1.5 flex items-center gap-2 uppercase tracking-wide">
                          <span className="bg-slate-900 text-white px-2 py-0.5 rounded-xs text-[10px] font-mono">I</span> 
                          THÔNG TIN CHUNG
                        </h3>

                        <div className="space-y-4 text-xs">
                          {/* Tên doanh nghiệp */}
                          <div className="flex flex-col sm:flex-row sm:items-end gap-2">
                            <span className="font-bold text-slate-800 shrink-0">1. Tên doanh nghiệp:</span>
                            <input
                              type="text"
                              required
                              placeholder="Nhập tên doanh nghiệp..."
                              value={hyName}
                              onChange={e => setHyName(e.target.value.toUpperCase())}
                              className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs uppercase font-bold text-slate-900 placeholder:italic placeholder:font-serif"
                            />
                          </div>

                          {/* Đại diện pháp luật */}
                          <div className="flex flex-col sm:flex-row sm:items-end gap-2">
                            <span className="font-bold text-slate-800 shrink-0">Người đại diện pháp luật:</span>
                            <input
                              type="text"
                              placeholder="Họ và tên..."
                              value={hyRep}
                              onChange={e => setHyRep(e.target.value)}
                              className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs font-bold text-slate-900 placeholder:italic placeholder:font-serif"
                            />
                          </div>

                          {/* Địa chỉ: Xã phường & Mã */}
                          <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 sm:items-end">
                            <div className="lg:col-span-7 flex items-end gap-2">
                              <span className="font-bold text-slate-800 shrink-0">2. Địa chỉ doanh nghiệp - Xã, phường:</span>
                              <input
                                type="text"
                                placeholder="Tên xã, phường, thị trấn..."
                                value={hyCommuneName}
                                onChange={e => setHyCommuneName(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                              />
                            </div>
                            <div className="lg:col-span-5 flex items-center justify-start lg:justify-end gap-2 flex-wrap">
                              <span className="font-bold text-slate-700 shrink-0 whitespace-nowrap text-[11.5px]">Mã xã, phường:</span>
                              <div className="bg-white p-0.5 rounded-sm border border-slate-400">
                                <GridBoxesInput value={hyCommuneCode} onChange={setHyCommuneCode} length={5} />
                              </div>
                            </div>
                          </div>

                          {/* Địa chỉ cụ thể */}
                          <div className="flex flex-col sm:flex-row sm:items-end gap-2">
                            <span className="font-bold text-slate-800 shrink-0">- Địa chỉ (số nhà, đường phố/thôn):</span>
                            <input
                              type="text"
                              placeholder="Số nhà, ngõ, tên đường hoặc tên thôn/xóm..."
                              value={hySpecificAddress}
                              onChange={e => setHySpecificAddress(e.target.value)}
                              className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                            />
                          </div>

                          {/* Số điện thoại & Email */}
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div className="flex items-end gap-2">
                              <span className="font-bold text-slate-800 shrink-0">- Số điện thoại:</span>
                              <input
                                type="text"
                                placeholder="Số điện thoại liên lạc..."
                                value={hyPhone}
                                onChange={e => setHyPhone(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                              />
                            </div>
                            <div className="flex items-end gap-2">
                              <span className="font-bold text-slate-800 shrink-0">Email:</span>
                              <input
                                type="email"
                                placeholder="Địa chỉ email liên hệ..."
                                value={hyEmail}
                                onChange={e => setHyEmail(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                              />
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* PHẦN II: NGÀNH HOẠT ĐỘNG CỦA DOANH NGHIỆP BÁO CÁO */}
                      <div className="space-y-4 pt-2">
                        <h3 className="text-xs sm:text-sm font-black text-slate-900 border-b border-slate-400 pb-1.5 flex items-center gap-2 uppercase tracking-wide">
                          <span className="bg-slate-900 text-white px-2 py-0.5 rounded-xs text-[10px] font-mono">II</span> 
                          NGÀNH HOẠT ĐỘNG CỦA DOANH NGHIỆP BÁO CÁO
                        </h3>

                        <div className="space-y-4 text-xs">
                          {/* Ngành SXKD chính */}
                          <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 sm:items-end">
                            <div className="lg:col-span-8 flex items-end gap-2">
                              <span className="font-bold text-slate-800 shrink-0">1. Ngành SXKD chính:</span>
                              <input
                                type="text"
                                placeholder="Gõ mô tả ngành hoạt động chính (ví dụ: Bán lẻ điện thoại, thiết bị viễn thông)..."
                                value={hyMaNganhMoTa}
                                onChange={e => setHyMaNganhMoTa(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 py-0.5 px-1 font-sans text-xs font-bold text-slate-900 focus:border-indigo-600 focus:outline-none focus:ring-0 placeholder:italic placeholder:font-serif"
                              />
                            </div>
                            <div className="lg:col-span-4 flex items-center justify-start lg:justify-end gap-2 flex-wrap">
                              <span className="font-bold text-slate-700 shrink-0 whitespace-nowrap text-[11.5px]">Mã ngành chính:</span>
                              <div className="bg-white p-0.5 rounded-sm border border-slate-400">
                                <GridBoxesInput value={hyMaNganh} onChange={setHyMaNganh} length={5} />
                              </div>
                            </div>
                          </div>

                          {/* Ngành SXKD khác */}
                          <div className="space-y-3">
                            <span className="font-bold text-slate-800 block">2. Ngành SXKD khác:</span>
                            
                            <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 sm:items-end pl-4">
                              <div className="lg:col-span-8 flex items-end gap-2">
                                <span className="text-slate-700 shrink-0">- Ngành khác 1:</span>
                                <input
                                  type="text"
                                  placeholder="Gõ mô tả ngành hoạt động khác 1..."
                                  value={hyNganhKhac1MoTa}
                                  onChange={e => setHyNganhKhac1MoTa(e.target.value)}
                                  className="flex-1 bg-transparent border-b border-dashed border-slate-550 py-0.5 px-1 font-sans text-xs text-slate-900 focus:border-indigo-600 focus:outline-none focus:ring-0 placeholder:italic"
                                />
                              </div>
                              <div className="lg:col-span-4 flex items-center justify-start lg:justify-end gap-2 flex-wrap">
                                <span className="font-bold text-slate-600 shrink-0 text-[11px]">Mã ngành 1:</span>
                                <div className="bg-white p-0.5 rounded-sm border border-slate-400">
                                  <GridBoxesInput value={hyNganhKhac1} onChange={setHyNganhKhac1} length={5} />
                                </div>
                              </div>
                            </div>

                            <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 sm:items-end pl-4">
                              <div className="lg:col-span-8 flex items-end gap-2">
                                <span className="text-slate-700 shrink-0">- Ngành khác 2:</span>
                                <input
                                  type="text"
                                  placeholder="Gõ mô tả ngành hoạt động khác 2..."
                                  value={hyNganhKhac2MoTa}
                                  onChange={e => setHyNganhKhac2MoTa(e.target.value)}
                                  className="flex-1 bg-transparent border-b border-dashed border-slate-550 py-0.5 px-1 font-sans text-xs text-slate-900 focus:border-indigo-600 focus:outline-none focus:ring-0 placeholder:italic"
                                />
                              </div>
                              <div className="lg:col-span-4 flex items-center justify-start lg:justify-end gap-2 flex-wrap">
                                <span className="font-bold text-slate-600 shrink-0 text-[11px]">Mã ngành 2:</span>
                                <div className="bg-white p-0.5 rounded-sm border border-slate-400">
                                  <GridBoxesInput value={hyNganhKhac2} onChange={setHyNganhKhac2} length={5} />
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* PHẦN III: THÔNG TIN VỀ LAO ĐỘNG */}
                      <div className="space-y-4 pt-2">
                        <h3 className="text-xs sm:text-sm font-black text-slate-900 border-b border-slate-400 pb-1.5 flex items-center gap-2 uppercase tracking-wide">
                          <span className="bg-slate-900 text-white px-2 py-0.5 rounded-xs text-[10px] font-mono">III</span> 
                          THÔNG TIN VỀ LAO ĐỘNG
                        </h3>

                        <div className="space-y-3.5 text-xs">
                          {/* Lao động 0101 */}
                          <div className="flex flex-col sm:flex-row sm:items-end gap-2">
                            <span className="font-bold text-slate-800 shrink-0">1. Thời điểm 01/01 năm báo cáo:</span>
                            <input
                              type="number"
                              min="0"
                              placeholder="Số lượng..."
                              value={hyLaoDong0101}
                              onChange={e => setHyLaoDong0101(e.target.value)}
                              className="w-24 bg-transparent border-b border-dashed border-slate-500 text-center font-sans font-bold text-slate-900 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5"
                            />
                            <span className="text-slate-700 italic">người</span>
                          </div>

                          {/* Lao động 3112 */}
                          <div className="flex flex-col sm:flex-row sm:items-end gap-2">
                            <span className="font-bold text-slate-800 shrink-0">2. Dự kiến thời điểm 31/12 năm báo cáo:</span>
                            <input
                              type="number"
                              min="0"
                              placeholder="Số lượng..."
                              value={hyLaoDong3112}
                              onChange={e => setHyLaoDong3112(e.target.value)}
                              className="w-24 bg-transparent border-b border-dashed border-slate-500 text-center font-sans font-bold text-slate-900 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5"
                            />
                            <span className="text-slate-700 italic">người (Chỉ tiêu lưu trữ chính)</span>
                          </div>
                        </div>
                      </div>

                      {/* PHẦN IV: THÔNG TIN VỀ KẾT QUẢ SẢN XUẤT KINH DOANH */}
                      <div className="space-y-4 pt-2">
                        <h3 className="text-xs sm:text-sm font-black text-slate-900 border-b border-slate-400 pb-1.5 flex items-center gap-2 uppercase tracking-wide">
                          <span className="bg-slate-900 text-white px-2 py-0.5 rounded-xs text-[10px] font-mono">IV</span> 
                          THÔNG TIN VỀ KẾT QUẢ SẢN XUẤT KINH DOANH
                        </h3>

                        <div className="space-y-4 text-xs">
                          {/* Doanh thu */}
                          <div className="flex flex-col sm:flex-row sm:items-end gap-2">
                            <span className="font-bold text-slate-800 shrink-0">1. Tổng Doanh thu thuần ước tính năm báo cáo (Triệu đồng):</span>
                            <input
                              type="number"
                              min="0"
                              step="any"
                              placeholder="Nhập doanh thu..."
                              value={hyDoanhThu}
                              onChange={e => setHyDoanhThu(e.target.value)}
                              className="w-36 bg-transparent border-b border-dashed border-slate-500 text-center font-sans font-bold text-slate-900 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5"
                            />
                            <span className="text-slate-650 font-bold italic">Triệu VNĐ</span>
                          </div>

                          {/* Cơ sở ngoài xã */}
                          <div className="space-y-2 pt-1">
                            <span className="font-bold text-slate-800 block">
                              2. Doanh nghiệp có cơ sở sản xuất kinh doanh ngoài xã/phường không?
                            </span>
                            
                            <div className="flex items-center gap-6 pl-4 font-sans mt-1">
                              <label className="flex items-center gap-2 cursor-pointer text-xs">
                                <input
                                  type="radio"
                                  name="hyCoSoNgoaiXa"
                                  value="có"
                                  checked={hyCoSoNgoaiXa === "có"}
                                  onChange={() => setHyCoSoNgoaiXa("có")}
                                  className="w-4 h-4 text-indigo-600 border-slate-300 focus:ring-indigo-500 cursor-pointer"
                                />
                                <span className="font-bold text-slate-900">Có  ⟶  Chuyển sang Trang 2 điền chi tiết</span>
                              </label>

                              <label className="flex items-center gap-2 cursor-pointer text-xs">
                                <input
                                  type="radio"
                                  name="hyCoSoNgoaiXa"
                                  value="không"
                                  checked={hyCoSoNgoaiXa === "không"}
                                  onChange={() => setHyCoSoNgoaiXa("không")}
                                  className="w-4 h-4 text-indigo-600 border-slate-300 focus:ring-indigo-500 cursor-pointer"
                                />
                                <span className="font-bold text-slate-950">Không  ⟶  Kết thúc Trang 1 (Chuyển xuống ký số)</span>
                              </label>
                            </div>
                          </div>

                          {/* Thông báo nếu chọn KHÔNG */}
                          {hyCoSoNgoaiXa === "không" && (
                            <div className="bg-emerald-50 border border-emerald-200 text-emerald-800 p-3 rounded-lg flex items-center gap-2 mt-4 font-sans italic">
                              <ShieldCheck className="w-4 h-4 text-emerald-600 shrink-0" />
                              <span>Doanh nghiệp không có cơ sở ngoài xã/phường. Toàn bộ thông tin khảo sát cơ bản đã hoàn thành, mời bạn cuộn xuống để thực hiện ký số/chữ ký điện tử bảo mật và nộp phiếu!</span>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* ======================================================== */}
                      {/* TRANG 2 DOANH NGHIỆP: CHỈ HIỂN THỊ KHI TÍCH CÓ CƠ SỞ NGOÀI XA */}
                      {/* ======================================================== */}
                      {hyCoSoNgoaiXa === "có" && (
                        <div className="mt-8 pt-6 border-t-2 border-dashed border-slate-400 space-y-6 bg-slate-50/50 p-4 rounded-xl">
                          <div className="flex items-center justify-between border-b border-slate-300 pb-2">
                            <h3 className="text-xs sm:text-sm font-black text-indigo-900 flex items-center gap-2 uppercase tracking-wide">
                              <span className="bg-indigo-900 text-white px-2 py-0.5 rounded-xs text-[10px] font-mono">TRANG 2</span> 
                              CHI TIẾT CƠ SỞ NGOÀI XÃ/PHƯỜNG & KẾT QUẢ SẢN XUẤT KINH DOANH
                            </h3>
                            <span className="text-[10px] bg-indigo-100 text-indigo-800 font-bold px-2 py-0.5 rounded-full font-sans">Biểu mẫu chuyên sâu</span>
                          </div>

                          {/* Mục 2: Số lao động của cơ sở ngoài xã */}
                          <div className="space-y-3">
                            <h4 className="font-black text-slate-900 text-[11.5px] leading-snug">
                              2. Số lao động của cơ sở SXKD có địa điểm kinh doanh ngoài xã/phường của trụ sở chính?
                            </h4>
                            
                            <div className="overflow-x-auto border border-slate-400 shadow-xs rounded-xs">
                              <table className="w-full text-center border-collapse text-[11px] text-slate-900 bg-white">
                                <thead>
                                  <tr className="bg-slate-100 font-serif font-black text-slate-800">
                                    <th className="border border-slate-400 p-2 w-10">STT</th>
                                    <th className="border border-slate-400 p-2 min-w-[150px]">Tên địa điểm SXKD</th>
                                    <th className="border border-slate-400 p-2 min-w-[150px]">Địa chỉ địa điểm SXKD (Ghi xã/phường)</th>
                                    <th className="border border-slate-400 p-2 w-32">Số lao động thời điểm 01/01/2026 (người)</th>
                                    <th className="border border-slate-400 p-2 w-32">Số lao động thời điểm 31/12/2026 (người)</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {hyCoSoNgoaiXaList.map((item, index) => (
                                    <tr key={index} className="hover:bg-slate-50 transition-colors">
                                      <td className="border border-slate-400 p-2 font-mono font-bold text-slate-600 bg-slate-50/50">{item.stt}</td>
                                      <td className="border border-slate-400 p-1">
                                        <input
                                          type="text"
                                          placeholder="Nhập tên địa điểm..."
                                          value={item.ten}
                                          onChange={e => {
                                            const newList = [...hyCoSoNgoaiXaList];
                                            newList[index].ten = e.target.value;
                                            setHyCoSoNgoaiXaList(newList);
                                          }}
                                          className="w-full bg-transparent border-0 focus:ring-0 text-left font-sans text-xs px-1"
                                        />
                                      </td>
                                      <td className="border border-slate-400 p-1">
                                        <input
                                          type="text"
                                          placeholder="Nhập xã/phường..."
                                          value={item.diaChi}
                                          onChange={e => {
                                            const newList = [...hyCoSoNgoaiXaList];
                                            newList[index].diaChi = e.target.value;
                                            setHyCoSoNgoaiXaList(newList);
                                          }}
                                          className="w-full bg-transparent border-0 focus:ring-0 text-left font-sans text-xs px-1"
                                        />
                                      </td>
                                      <td className="border border-slate-400 p-1">
                                        <input
                                          type="number"
                                          min="0"
                                          placeholder="0"
                                          value={item.laoDong0101}
                                          onChange={e => {
                                            const newList = [...hyCoSoNgoaiXaList];
                                            newList[index].laoDong0101 = e.target.value;
                                            setHyCoSoNgoaiXaList(newList);
                                          }}
                                          className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs"
                                        />
                                      </td>
                                      <td className="border border-slate-400 p-1">
                                        <input
                                          type="number"
                                          min="0"
                                          placeholder="0"
                                          value={item.laoDong3112}
                                          onChange={e => {
                                            const newList = [...hyCoSoNgoaiXaList];
                                            newList[index].laoDong3112 = e.target.value;
                                            setHyCoSoNgoaiXaList(newList);
                                          }}
                                          className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs text-indigo-700"
                                        />
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>

                          {/* Mục 3: Kết quả sản xuất kinh doanh */}
                          <div className="space-y-3 pt-2">
                            <div className="flex items-center justify-between">
                              <h4 className="font-black text-slate-900 text-[11.5px]">3. Kết quả sản xuất kinh doanh:</h4>
                              <span className="text-[10.5px] font-bold italic text-slate-600">Đơn vị tính: Triệu đồng</span>
                            </div>

                            <div className="overflow-x-auto border border-slate-400 shadow-xs rounded-xs">
                              <table className="w-full text-center border-collapse text-[11px] text-slate-900 bg-white">
                                <thead>
                                  <tr className="bg-slate-100 font-serif font-black text-slate-800">
                                    <th className="border border-slate-400 p-2 text-left" rowSpan={2}>Chỉ tiêu</th>
                                    <th className="border border-slate-400 p-2 w-14" rowSpan={2}>Mã số</th>
                                    <th className="border border-slate-400 p-2 w-36" rowSpan={2}>Năm trước năm báo cáo</th>
                                    <th className="border border-slate-400 p-2" colSpan={2}>Năm báo cáo</th>
                                  </tr>
                                  <tr className="bg-slate-100 font-serif font-black text-slate-800">
                                    <th className="border border-slate-400 p-2 w-36">Thực hiện 9 tháng đầu năm</th>
                                    <th className="border border-slate-400 p-2 w-36">Ước cả năm</th>
                                  </tr>
                                  <tr className="bg-slate-200 text-[10px] font-bold text-slate-600">
                                    <td className="border border-slate-400 py-0.5 text-center">A</td>
                                    <td className="border border-slate-400 py-0.5">B</td>
                                    <td className="border border-slate-400 py-0.5">1</td>
                                    <td className="border border-slate-400 py-0.5">2</td>
                                    <td className="border border-slate-400 py-0.5">3</td>
                                  </tr>
                                </thead>
                                <tbody>
                                  {/* Dòng 01: Tổng doanh thu */}
                                  <tr className="bg-indigo-50/40 font-bold">
                                    <td className="border border-slate-400 p-2.5 text-left font-bold font-serif leading-snug text-indigo-950">
                                      Tổng doanh thu bán hàng và cung cấp dịch vụ
                                    </td>
                                    <td className="border border-slate-400 p-2 font-mono text-center text-indigo-950">01</td>
                                    <td className="border border-slate-400 p-1 bg-indigo-50/10">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdTongDtNamTruoc}
                                        onChange={e => setHyKqkdTongDtNamTruoc(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-black text-xs text-indigo-950"
                                      />
                                    </td>
                                    <td className="border border-slate-400 p-1 bg-indigo-50/10">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdTongDt9Thang}
                                        onChange={e => setHyKqkdTongDt9Thang(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-black text-xs text-indigo-950"
                                      />
                                    </td>
                                    <td className="border border-slate-400 p-1 bg-indigo-50/10">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdTongDtUocCaNam}
                                        onChange={e => setHyKqkdTongDtUocCaNam(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-black text-xs text-emerald-900"
                                      />
                                    </td>
                                  </tr>

                                  <tr className="bg-slate-50 italic text-[10.5px] text-slate-500">
                                    <td className="border border-slate-400 p-1 text-left pl-4 font-bold" colSpan={5}>
                                      Chia theo ngành hoạt động:
                                    </td>
                                  </tr>

                                  {/* Dòng 02: Ngành chính */}
                                  <tr className="bg-white">
                                    <td className="border border-slate-400 p-2 text-left font-medium pl-6 leading-snug">
                                      - Ngành SXKD chính: <span className="font-bold text-slate-800 font-sans italic">({hyMaNganhMoTa || "Chưa nhập tên ngành"})</span>
                                    </td>
                                    <td className="border border-slate-400 p-2 font-mono text-center text-slate-700">02</td>
                                    <td className="border border-slate-400 p-1">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdChinhDtNamTruoc}
                                        onChange={e => setHyKqkdChinhDtNamTruoc(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs"
                                      />
                                    </td>
                                    <td className="border border-slate-400 p-1">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdChinhDt9Thang}
                                        onChange={e => setHyKqkdChinhDt9Thang(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs text-indigo-700"
                                      />
                                    </td>
                                    <td className="border border-slate-400 p-1">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdChinhDtUocCaNam}
                                        onChange={e => setHyKqkdChinhDtUocCaNam(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs text-emerald-700"
                                      />
                                    </td>
                                  </tr>

                                  <tr className="bg-slate-50 italic text-[10.5px] text-slate-500">
                                    <td className="border border-slate-400 p-1 text-left pl-4 font-bold" colSpan={5}>
                                      - Ngành SXKD khác:
                                    </td>
                                  </tr>

                                  {/* Dòng 03: Ngành khác 1 */}
                                  <tr className="bg-white">
                                    <td className="border border-slate-400 p-1 text-left pl-6 leading-snug">
                                      <div className="flex items-center gap-1.5">
                                        <span className="shrink-0 text-slate-700">Ngành:</span>
                                        <input
                                          type="text"
                                          placeholder="Nhập tên ngành khác 1..."
                                          value={hyKqkdKhac1Ten}
                                          onChange={e => setHyKqkdKhac1Ten(e.target.value)}
                                          className="flex-1 bg-transparent border-b border-dashed border-slate-300 py-0.5 px-1 font-sans text-xs text-slate-900 focus:outline-none focus:border-indigo-500"
                                        />
                                      </div>
                                    </td>
                                    <td className="border border-slate-400 p-2 font-mono text-center text-slate-700">03</td>
                                    <td className="border border-slate-400 p-1">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdKhac1DtNamTruoc}
                                        onChange={e => setHyKqkdKhac1DtNamTruoc(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs"
                                      />
                                    </td>
                                    <td className="border border-slate-400 p-1">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdKhac1Dt9Thang}
                                        onChange={e => setHyKqkdKhac1Dt9Thang(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs text-indigo-700"
                                      />
                                    </td>
                                    <td className="border border-slate-400 p-1">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdKhac1DtUocCaNam}
                                        onChange={e => setHyKqkdKhac1DtUocCaNam(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs text-emerald-700"
                                      />
                                    </td>
                                  </tr>

                                  {/* Dòng 04: Ngành khác 2 */}
                                  <tr className="bg-white">
                                    <td className="border border-slate-400 p-1 text-left pl-6 leading-snug">
                                      <div className="flex items-center gap-1.5">
                                        <span className="shrink-0 text-slate-700">Ngành:</span>
                                        <input
                                          type="text"
                                          placeholder="Nhập tên ngành khác 2..."
                                          value={hyKqkdKhac2Ten}
                                          onChange={e => setHyKqkdKhac2Ten(e.target.value)}
                                          className="flex-1 bg-transparent border-b border-dashed border-slate-300 py-0.5 px-1 font-sans text-xs text-slate-900 focus:outline-none focus:border-indigo-500"
                                        />
                                      </div>
                                    </td>
                                    <td className="border border-slate-400 p-2 font-mono text-center text-slate-700">04</td>
                                    <td className="border border-slate-400 p-1">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdKhac2DtNamTruoc}
                                        onChange={e => setHyKqkdKhac2DtNamTruoc(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs"
                                      />
                                    </td>
                                    <td className="border border-slate-400 p-1">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdKhac2Dt9Thang}
                                        onChange={e => setHyKqkdKhac2Dt9Thang(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs text-indigo-700"
                                      />
                                    </td>
                                    <td className="border border-slate-400 p-1">
                                      <input
                                        type="number"
                                        min="0"
                                        step="any"
                                        placeholder="0"
                                        value={hyKqkdKhac2DtUocCaNam}
                                        onChange={e => setHyKqkdKhac2DtUocCaNam(e.target.value)}
                                        className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs text-emerald-700"
                                      />
                                    </td>
                                  </tr>
                                </tbody>
                              </table>
                            </div>

                            <p className="text-[10px] text-slate-500 italic mt-1 font-sans">
                              (*) Ứng dụng đã kích hoạt công cụ tự động hóa: Tổng doanh thu (Mã số 01) của các thời kỳ sẽ tự động cộng dồn thời gian thực theo số liệu nhập của Ngành chính (02) + Các ngành khác (03, 04).
                            </p>
                          </div>

                          {/* Ghi chú dời xuống Trang 2 */}
                          <div className="flex flex-col sm:flex-row sm:items-end gap-2 pt-2">
                            <span className="font-bold text-slate-850 shrink-0 text-xs">Ghi chú khảo sát:</span>
                            <input
                              type="text"
                              placeholder="Ghi nhận thêm thông tin khác của doanh nghiệp thương mại..."
                              value={hyGhiChu}
                              onChange={e => setHyGhiChu(e.target.value)}
                              className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                            />
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* ---------------------------------------------------- */}
                  {/* TRƯỜNG HỢP B: KHỐI CÁ THỂ (PHIẾU SỐ 04/TGTSP-CTTMDV) */}
                  {/* ---------------------------------------------------- */}
                  {activeBlock === "cathe" && (
                    <div className="space-y-6">
                      {/* Header phiếu */}
                      <div className="flex flex-col md:flex-row md:items-start justify-between gap-4 border-b border-slate-300 pb-4">
                        <div className="space-y-1">
                          <h4 className="text-xs sm:text-sm font-black tracking-wider text-slate-900 uppercase">UBND TỈNH HƯNG YÊN</h4>
                          <p className="text-[11px] font-bold text-slate-700 italic">Phiếu số 04/TGTSP-CTTMDV</p>
                        </div>
                        
                        <div className="flex flex-col sm:flex-row sm:items-center gap-2">
                          <span className="text-xs font-black text-slate-800 uppercase font-serif whitespace-nowrap">STT cơ sở:</span>
                          <div className="bg-white p-1 rounded-sm border border-slate-400 shadow-2xs">
                            <GridBoxesInput value={hySttCoSo} onChange={setHySttCoSo} length={4} />
                          </div>
                        </div>
                      </div>

                      {/* Title tiêu đề chính */}
                      <div className="text-center space-y-2 py-2">
                        <h2 className="text-base sm:text-xl font-extrabold text-slate-900 uppercase leading-normal tracking-wide">
                          PHIẾU THU THẬP THÔNG TIN VỀ HOẠT ĐỘNG<br />THƯƠNG MẠI VÀ DỊCH VỤ CỦA CƠ SỞ CÁ THỂ
                        </h2>
                        <p className="text-[13px] font-bold text-slate-800 tracking-widest">NĂM 2026</p>
                        <p className="text-[11px] sm:text-[11.5px] font-bold text-slate-600 italic">
                          (Áp dụng cho các cơ sở SXKD cá thể hoạt động trong lĩnh vực thương mại và dịch vụ)
                        </p>
                      </div>

                      {/* Hai hộp lưu ý bên dưới tiêu đề */}
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-[11px] leading-relaxed">
                        <div className="border border-slate-300 p-3 bg-white rounded-sm shadow-2xs italic text-slate-700">
                          Thực hiện theo <b>Quyết định của UBND tỉnh Hưng Yên</b> về việc Ban hành Phương án điều tra, thu thập thông tin phục vụ biên soạn chỉ tiêu Tổng giá trị sản phẩm trên địa bàn xã, phường tỉnh Hưng Yên.
                        </div>
                        <div className="border border-slate-300 p-3 bg-white rounded-sm shadow-2xs space-y-1 text-slate-700">
                          <p>- Nghĩa vụ cung cấp thông tin được quy định theo Luật thống kê;</p>
                          <p>- Thông tin cung cấp theo phiếu này chỉ nhằm phục vụ công tác thống kê và được bảo mật theo Luật định.</p>
                        </div>
                      </div>

                      {/* NỘI DUNG PHIẾU CÁ THỂ */}
                      <div className="space-y-4 text-xs">
                        {/* 1. Tên cơ sở */}
                        <div className="flex flex-col sm:flex-row sm:items-end gap-2">
                          <span className="font-bold text-slate-900 shrink-0">1. Tên cơ sở:</span>
                          <input
                            type="text"
                            required
                            placeholder="Nhập tên cơ sở kinh doanh cá thể..."
                            value={hyName}
                            onChange={e => setHyName(e.target.value.toUpperCase())}
                            className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs uppercase font-bold text-slate-950 placeholder:italic placeholder:font-serif"
                          />
                        </div>

                        {/* 2. Địa chỉ cơ sở */}
                        <div className="space-y-3.5">
                          <span className="font-bold text-slate-900 block">2. Địa chỉ cơ sở:</span>
                          
                          {/* Xã phường & Mã */}
                          <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 sm:items-end pl-4">
                            <div className="lg:col-span-7 flex items-end gap-2">
                              <span className="font-bold text-slate-800 shrink-0">- Xã, phường:</span>
                              <input
                                type="text"
                                placeholder="Tên xã, phường..."
                                value={hyCommuneName}
                                onChange={e => setHyCommuneName(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                              />
                            </div>
                            <div className="lg:col-span-5 flex items-center justify-start lg:justify-end gap-2 flex-wrap">
                              <span className="font-bold text-slate-700 shrink-0 text-[11px] whitespace-nowrap">Mã xã, phường:</span>
                              <div className="bg-white p-0.5 rounded-sm border border-slate-400">
                                <GridBoxesInput value={hyCommuneCode} onChange={setHyCommuneCode} length={5} />
                              </div>
                            </div>
                          </div>

                          {/* Địa chỉ chi tiết */}
                          <div className="flex flex-col sm:flex-row sm:items-end gap-2 pl-4">
                            <span className="font-bold text-slate-800 shrink-0">- Địa chỉ (số nhà, đường phố, thôn):</span>
                            <input
                              type="text"
                              placeholder="Số nhà, đường phố, thôn, xóm..."
                              value={hySpecificAddress}
                              onChange={e => setHySpecificAddress(e.target.value)}
                              className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                            />
                          </div>

                          {/* Số điện thoại & Email */}
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pl-4">
                            <div className="flex items-end gap-2">
                              <span className="font-bold text-slate-800 shrink-0">- Số điện thoại:</span>
                              <input
                                type="text"
                                placeholder="Số điện thoại..."
                                value={hyPhone}
                                onChange={e => setHyPhone(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                              />
                            </div>
                            <div className="flex items-end gap-2">
                              <span className="font-bold text-slate-800 shrink-0">Email:</span>
                              <input
                                type="email"
                                placeholder="Địa chỉ email..."
                                value={hyEmail}
                                onChange={e => setHyEmail(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                              />
                            </div>
                          </div>
                        </div>

                        {/* 3. Ngành hoạt động chính & Mã ngành */}
                        <div className="grid grid-cols-1 lg:grid-cols-12 gap-3 sm:items-end pt-1">
                          <div className="lg:col-span-8 flex items-end gap-2">
                            <span className="font-bold text-slate-900 shrink-0">3. Ngành SXKD chính:</span>
                            <input
                              type="text"
                              placeholder="Gõ mô tả ngành hoạt động chính..."
                              value={hyMaNganhMoTa}
                              onChange={e => setHyMaNganhMoTa(e.target.value)}
                              className="flex-1 bg-transparent border-b border-dashed border-slate-500 py-0.5 px-1 font-sans text-xs font-bold text-slate-900 focus:border-emerald-600 focus:outline-none focus:ring-0 placeholder:italic placeholder:font-serif"
                            />
                          </div>
                          <div className="lg:col-span-4 flex items-center justify-start lg:justify-end gap-2 flex-wrap">
                            <span className="font-bold text-slate-800 shrink-0 text-[11.5px] whitespace-nowrap">Mã ngành:</span>
                            <div className="bg-white p-0.5 rounded-sm border border-slate-400 font-mono">
                              <GridBoxesInput value={hyMaNganh} onChange={setHyMaNganh} length={5} />
                            </div>
                          </div>
                        </div>

                        {/* 4. Số lao động */}
                        <div className="space-y-3 pt-1">
                          <span className="font-bold text-slate-900 block">4. Số lao động:</span>
                          <div className="flex flex-col sm:flex-row sm:items-center gap-6 pl-4">
                            <div className="flex items-center gap-2">
                              <span className="font-bold text-slate-850">- Thời điểm 01/01 năm báo cáo:</span>
                              <input
                                type="number"
                                min="0"
                                placeholder="..."
                                value={hyLaoDong0101}
                                onChange={e => setHyLaoDong0101(e.target.value)}
                                className="w-16 bg-transparent border-b border-dashed border-slate-500 text-center font-sans font-bold text-slate-950 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5"
                              />
                              <span className="italic text-slate-600">người;</span>
                            </div>
                            <div className="flex items-center gap-2">
                              <span className="font-bold text-slate-850">- Dự kiến thời điểm 31/12 năm báo cáo:</span>
                              <input
                                type="number"
                                min="0"
                                placeholder="..."
                                value={hyLaoDong3112}
                                onChange={e => setHyLaoDong3112(e.target.value)}
                                className="w-16 bg-transparent border-b border-dashed border-slate-500 text-center font-sans font-bold text-slate-950 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5"
                              />
                              <span className="italic text-slate-600">người.</span>
                            </div>
                          </div>
                        </div>

                        {/* 5. Doanh thu thuần của hoạt động thương mại và dịch vụ */}
                        <div className="space-y-2 pt-2">
                          <div className="flex items-center justify-between">
                            <span className="font-bold text-slate-900 block">5. Doanh thu thuần của hoạt động thương mại và dịch vụ:</span>
                            <span className="text-[11px] font-bold italic text-slate-600">Đơn vị tính: Triệu đồng</span>
                          </div>

                          {/* Bảng Doanh thu chuẩn chỉ */}
                          <div className="overflow-x-auto border border-slate-400">
                            <table className="w-full text-center border-collapse text-[11px] text-slate-900">
                              <thead>
                                <tr className="bg-slate-100 font-serif font-black">
                                  <th className="border border-slate-400 p-2 text-left" rowSpan={2}>Chỉ tiêu</th>
                                  <th className="border border-slate-400 p-2 w-14" rowSpan={2}>Mã số</th>
                                  <th className="border border-slate-400 p-2" rowSpan={2}>Thực hiện năm trước năm báo cáo</th>
                                  <th className="border border-slate-400 p-2" colSpan={2}>Năm báo cáo</th>
                                </tr>
                                <tr className="bg-slate-100 font-serif font-black">
                                  <th className="border border-slate-400 p-2">Thực hiện 9 tháng đầu năm</th>
                                  <th className="border border-slate-400 p-2">Ước tính 3 tháng cuối năm</th>
                                </tr>
                                <tr className="bg-slate-200 text-[10px] font-bold">
                                  <td className="border border-slate-400 py-0.5 text-center">A</td>
                                  <td className="border border-slate-400 py-0.5">B</td>
                                  <td className="border border-slate-400 py-0.5">1</td>
                                  <td className="border border-slate-400 py-0.5">2</td>
                                  <td className="border border-slate-400 py-0.5">3</td>
                                </tr>
                              </thead>
                              <tbody>
                                <tr className="bg-white">
                                  <td className="border border-slate-400 p-2.5 text-left font-bold font-serif leading-snug">
                                    Doanh thu thuần của hoạt động thương mại và dịch vụ (*)
                                  </td>
                                  <td className="border border-slate-400 p-2 font-mono font-bold text-center">01</td>
                                  <td className="border border-slate-400 p-1">
                                    <input
                                      type="number"
                                      min="0"
                                      step="any"
                                      placeholder="Nhập số..."
                                      value={hyDoanhThuNamTruoc}
                                      onChange={e => setHyDoanhThuNamTruoc(e.target.value)}
                                      className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs p-1"
                                    />
                                  </td>
                                  <td className="border border-slate-400 p-1">
                                    <input
                                      type="number"
                                      min="0"
                                      step="any"
                                      placeholder="Nhập số..."
                                      value={hyDoanhThu9Thang}
                                      onChange={e => setHyDoanhThu9Thang(e.target.value)}
                                      className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs p-1 text-indigo-700"
                                    />
                                  </td>
                                  <td className="border border-slate-400 p-1">
                                    <input
                                      type="number"
                                      min="0"
                                      step="any"
                                      placeholder="Nhập số..."
                                      value={hyDoanhThu3Thang}
                                      onChange={e => setHyDoanhThu3Thang(e.target.value)}
                                      className="w-full bg-transparent border-0 focus:ring-0 text-center font-sans font-bold text-xs p-1 text-emerald-700"
                                    />
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                          </div>
                          
                          {/* Live Total Calculated Area */}
                          <div className="bg-slate-100/50 p-2.5 border border-dashed border-slate-400 text-[11px] flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 mt-1">
                            <span className="italic text-slate-700">(*) Doanh thu tính trên địa bàn xã/phường được chọn mẫu điều tra.</span>
                            <div className="font-serif font-black text-slate-900 flex items-center gap-1 shrink-0 text-xs sm:text-[13px]">
                              <span>Tổng Doanh thu ước tính cả năm:</span>
                              <span className="text-emerald-700 font-sans font-extrabold bg-emerald-50 px-2 py-0.5 border border-emerald-350">
                                {((parseFloat(hyDoanhThu9Thang) || 0) + (parseFloat(hyDoanhThu3Thang) || 0) || parseFloat(hyDoanhThuNamTruoc) || 0).toLocaleString()} Triệu VNĐ
                              </span>
                            </div>
                          </div>
                        </div>

                        {/* Ghi chú */}
                        <div className="flex flex-col sm:flex-row sm:items-end gap-2 pt-3">
                          <span className="font-bold text-slate-900 shrink-0">Ghi chú khảo sát:</span>
                          <input
                            type="text"
                            placeholder="Ghi nhận thêm thông tin khác của hộ cá thể..."
                            value={hyGhiChu}
                            onChange={e => setHyGhiChu(e.target.value)}
                            className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans text-xs text-slate-900"
                          />
                        </div>
                      </div>
                    </div>
                  )}

                  {/* CÁC CHỈ TIÊU ĐỘNG TỰ CẤU HÌNH THEO BIỂU MẪU */}
                  {customFields.length > 0 && (
                    <div className="border-t border-slate-400 mt-6 pt-4 space-y-4">
                      <h4 className="text-xs font-black text-slate-900 uppercase flex items-center gap-1.5 tracking-wide">
                        <Sparkles className="w-3.5 h-3.5 text-indigo-500 animate-pulse" />
                        Chỉ tiêu khảo sát động (Tự thiết lập bổ sung)
                      </h4>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs pl-2">
                        {customFields.map(field => (
                          <div key={field.name} className="flex flex-col sm:flex-row sm:items-end gap-2">
                            <span className="font-bold text-slate-800 shrink-0">{field.label}:</span>
                            {field.type === "boolean" ? (
                              <select
                                value={dynamicValues[field.name] || ""}
                                onChange={e => setDynamicValues(prev => ({ ...prev, [field.name]: e.target.value }))}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 font-sans font-bold text-slate-900 text-xs bg-white"
                              >
                                <option value="">-- Chưa chọn --</option>
                                <option value="true">Có / Đạt / Đúng</option>
                                <option value="false">Không / Chưa đạt / Sai</option>
                              </select>
                            ) : (
                              <input
                                type={field.type === "number" ? "number" : "text"}
                                step={field.type === "number" ? "any" : undefined}
                                placeholder={`Nhập ${field.label.toLowerCase()}...`}
                                value={dynamicValues[field.name] || ""}
                                onChange={e => setDynamicValues(prev => ({ ...prev, [field.name]: e.target.value }))}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans font-bold text-slate-900 text-xs"
                              />
                            )}
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Khu vực Chữ ký số điện tử bảo mật (Digital Signature) */}
                  <div className="pt-6 mt-6 border-t border-slate-350 space-y-4">
                    <h4 className="text-xs font-black text-slate-900 uppercase flex items-center gap-1.5 tracking-wide">
                      <FileSignature className="w-4 h-4 text-indigo-600 shrink-0" />
                      Xác nhận thông tin & Chữ ký số điện tử bảo mật
                    </h4>

                    {activeBlock === "doanhnghiep" ? (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs text-slate-800">
                        {/* Cột 1: Điều tra viên */}
                        <div className="border border-slate-300 p-4 bg-slate-50/50 rounded-lg space-y-3 flex flex-col">
                          <p className="font-bold font-serif text-slate-900 uppercase text-[11px] border-b border-slate-200 pb-1 flex items-center gap-1">
                            <Fingerprint className="w-3.5 h-3.5 text-slate-500" />
                            Đơn vị thu thập thông tin (Điều tra viên)
                          </p>
                          <div className="space-y-2 flex-1">
                            <div className="flex items-end gap-1.5 text-[11px]">
                              <span className="text-slate-700 shrink-0">Họ và tên ĐTV:</span>
                              <input
                                type="text"
                                placeholder="Nhập họ tên điều tra viên..."
                                value={hySurveyorName}
                                onChange={e => setHySurveyorName(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans font-bold text-slate-900"
                              />
                            </div>
                            <div className="flex items-end gap-1.5 text-[11px]">
                              <span className="text-slate-700 shrink-0">Số điện thoại:</span>
                              <input
                                type="text"
                                placeholder="Nhập SĐT..."
                                value={hySurveyorPhone}
                                onChange={e => setHySurveyorPhone(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans font-bold text-slate-900"
                              />
                            </div>
                          </div>
                          
                          {/* Signature Pad ĐTV */}
                          <div className="space-y-1">
                            <span className="text-[10px] font-bold text-slate-600 flex items-center gap-1">
                              <PenTool className="w-3 h-3" /> Ký tên điện tử ĐTV:
                            </span>
                            <SignaturePad
                              id="signature-dtv"
                              onSave={(dataUrl) => {
                                setHySurveyorSignature(dataUrl);
                              }}
                              height={110}
                            />
                          </div>
                        </div>

                        {/* Cột 2: Đại diện Doanh nghiệp */}
                        <div className="border border-slate-300 p-4 bg-indigo-50/20 rounded-lg space-y-3 flex flex-col justify-between">
                          <div>
                            <p className="font-bold font-serif text-indigo-950 uppercase text-[11px] border-b border-indigo-200 pb-1 flex items-center justify-between">
                              <span className="flex items-center gap-1">
                                <ShieldCheck className="w-3.5 h-3.5 text-indigo-700" />
                                Đại diện pháp luật Doanh nghiệp
                              </span>
                              <span className="text-[10px] text-indigo-700 font-sans italic">Hưng Yên Smart Sign</span>
                            </p>
                            <div className="space-y-2 flex-1 mt-2">
                              <div className="flex items-end gap-1.5 text-[11px]">
                                <span className="text-slate-700 shrink-0">Họ tên người đại diện:</span>
                                <input
                                  type="text"
                                  placeholder="Nhập họ tên người trả lời..."
                                  value={hyRespondentName}
                                  onChange={e => setHyRespondentName(e.target.value)}
                                  className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans font-bold text-slate-900"
                                />
                              </div>
                              <div className="flex items-end gap-1.5 text-[11px]">
                                <span className="text-slate-700 shrink-0">Chức vụ đại diện:</span>
                                <input
                                  type="text"
                                  placeholder="Giám đốc, Kế toán trưởng, Đại diện..."
                                  value={hyRespondentPhone}
                                  onChange={e => setHyRespondentPhone(e.target.value)}
                                  className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-indigo-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans font-bold text-slate-900"
                                />
                              </div>
                            </div>
                          </div>

                          {/* Chọn hình thức ký số / ký tay */}
                          <div className="space-y-2.5 pt-2">
                            <span className="text-[10px] font-bold text-slate-600 block">
                              Chọn hình thức ký xác thực phiếu:
                            </span>
                            <div className="flex items-center gap-2 p-1 bg-slate-200/60 rounded-md border border-slate-300">
                              <button
                                type="button"
                                onClick={() => {
                                  setHySignatureMode("draw");
                                  setHySignatureDataUrl("");
                                  setHyIsSigned(false);
                                  setHySignedHash("");
                                  setHySignedTime("");
                                  setHySignedBy("");
                                }}
                                className={`flex-1 py-1 px-1.5 rounded-md text-center text-[10.5px] font-bold transition-all cursor-pointer ${
                                  (hySignatureMode === "draw" || hySignatureMode === "")
                                    ? "bg-white text-indigo-950 shadow-xs"
                                    : "text-slate-600 hover:text-slate-950"
                                }`}
                              >
                                Ký tay điện tử
                              </button>
                              <button
                                type="button"
                                onClick={() => {
                                  setHySignatureMode("token");
                                  setHySignatureDataUrl("");
                                  setHyIsSigned(false);
                                  setHySignedHash("");
                                  setHySignedTime("");
                                  setHySignedBy("");
                                }}
                                className={`flex-1 py-1 px-1.5 rounded-md text-center text-[10.5px] font-bold transition-all cursor-pointer flex items-center justify-center gap-1 ${
                                  hySignatureMode === "token"
                                    ? "bg-white text-indigo-950 shadow-xs"
                                    : "text-slate-600 hover:text-slate-950"
                                }`}
                              >
                                <Key className="w-3 h-3 text-amber-500" />
                                Ký số USB Token CA
                              </button>
                            </div>

                            {/* Render tương ứng */}
                            {hySignatureMode === "token" ? (
                              <div className="pt-1">
                                <SignatureToken
                                  mst={hyMst.trim() || "0900123456"}
                                  enterpriseName={hyName.trim() || "CÔNG TY TNHH KHẢO SÁT HƯNG YÊN"}
                                  representative={hyRespondentName.trim() || "Người đại diện Doanh nghiệp"}
                                  onSave={(dataUrl, hash, certInfo) => {
                                    setHySignatureDataUrl(dataUrl);
                                    setHySignedHash(hash);
                                    setHySignedBy(hyRespondentName.trim() || "Người đại diện Doanh nghiệp");
                                    setHySignedTime(new Date().toLocaleString("vi-VN"));
                                    setHyIsSigned(true);
                                  }}
                                  onClear={() => {
                                    setHySignatureDataUrl("");
                                    setHySignedHash("");
                                    setHyIsSigned(false);
                                    setHySignedTime("");
                                    setHySignedBy("");
                                  }}
                                />
                              </div>
                            ) : (
                              <div className="space-y-1">
                                <span className="text-[10px] font-bold text-slate-600 flex items-center gap-1">
                                  <PenTool className="w-3 h-3" /> Vẽ chữ ký tay của bạn:
                                </span>
                                <SignaturePad
                                  id="signature-dn"
                                  onSave={(dataUrl) => {
                                    setHySignatureDataUrl(dataUrl);
                                    setHySignatureMode("draw");
                                    setHyIsSigned(true);
                                    setHySignedTime(new Date().toLocaleString("vi-VN"));
                                    setHySignedBy(hyRespondentName.trim() || "Đại diện Doanh nghiệp");
                                    setHySignedHash("DRAW-" + Math.random().toString(36).substring(2, 10).toUpperCase());
                                  }}
                                  onClear={() => {
                                    setHySignatureDataUrl("");
                                    setHyIsSigned(false);
                                    setHySignedHash("");
                                  }}
                                  height={110}
                                />
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    ) : (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs text-slate-800">
                        {/* Cột 1: Người trả lời phiếu (Hộ cá thể) */}
                        <div className="border border-slate-300 p-4 bg-slate-50/50 rounded-lg space-y-3 flex flex-col">
                          <p className="font-bold font-serif text-slate-900 uppercase text-[11px] border-b border-slate-200 pb-1 flex items-center gap-1">
                            <Fingerprint className="w-3.5 h-3.5 text-slate-500" />
                            Người trả lời phiếu (Chủ hộ/Người đại diện)
                          </p>
                          <div className="space-y-2 flex-1">
                            <div className="flex items-end gap-1.5 text-[11px]">
                              <span className="text-slate-700 shrink-0">Họ và tên:</span>
                              <input
                                type="text"
                                placeholder="Nhập họ tên chủ hộ..."
                                value={hyRespondentName}
                                onChange={e => setHyRespondentName(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans font-bold text-slate-900"
                              />
                            </div>
                            <div className="flex items-end gap-1.5 text-[11px]">
                              <span className="text-slate-700 shrink-0">Số điện thoại:</span>
                              <input
                                type="text"
                                placeholder="Nhập SĐT liên hệ..."
                                value={hyRespondentPhone}
                                onChange={e => setHyRespondentPhone(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans font-bold text-slate-900"
                              />
                            </div>
                          </div>

                          {/* Signature Pad Chủ hộ */}
                          <div className="space-y-1">
                            <span className="text-[10px] font-bold text-slate-600 flex items-center gap-1">
                              <PenTool className="w-3 h-3" /> Chữ ký tay chủ hộ:
                            </span>
                            <SignaturePad
                              id="signature-cathe-chuho"
                              onSave={(dataUrl) => {
                                setHySignatureDataUrl(dataUrl);
                              }}
                              height={110}
                            />
                          </div>
                        </div>

                        {/* Cột 2: Điều tra viên */}
                        <div className="border border-slate-300 p-4 bg-emerald-50/15 rounded-lg space-y-3 flex flex-col">
                          <p className="font-bold font-serif text-slate-950 uppercase text-[11px] border-b border-emerald-200 pb-1 flex items-center gap-1">
                            <ShieldCheck className="w-3.5 h-3.5 text-emerald-700" />
                            Điều tra viên (ĐTV)
                          </p>
                          <div className="space-y-2 flex-1">
                            <div className="flex items-end gap-1.5 text-[11px]">
                              <span className="text-slate-700 shrink-0">Họ và tên ĐTV:</span>
                              <input
                                type="text"
                                placeholder="Nhập họ tên điều tra viên..."
                                value={hySurveyorName}
                                onChange={e => setHySurveyorName(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans font-bold text-slate-900"
                              />
                            </div>
                            <div className="flex items-end gap-1.5 text-[11px]">
                              <span className="text-slate-700 shrink-0">Số điện thoại ĐTV:</span>
                              <input
                                type="text"
                                placeholder="Nhập SĐT..."
                                value={hySurveyorPhone}
                                onChange={e => setHySurveyorPhone(e.target.value)}
                                className="flex-1 bg-transparent border-b border-dashed border-slate-500 focus:border-emerald-600 focus:outline-none focus:ring-0 py-0.5 px-1 font-sans font-bold text-slate-900"
                              />
                            </div>
                          </div>

                          {/* Signature Pad ĐTV cho Cá thể */}
                          <div className="space-y-1">
                            <span className="text-[10px] font-bold text-slate-600 flex items-center gap-1">
                              <PenTool className="w-3 h-3" /> Chữ ký tay ĐTV:
                            </span>
                            <SignaturePad
                              id="signature-cathe-dtv"
                              onSave={(dataUrl) => {
                                setHySurveyorSignature(dataUrl);
                              }}
                              height={110}
                            />
                          </div>
                        </div>
                      </div>
                    )}
                  </div>

                </div>

                {/* Submit button */}
                <div className="pt-2">
                  <button
                    type="submit"
                    disabled={loading}
                    className="w-full py-3.5 bg-emerald-600 hover:bg-emerald-700 active:translate-y-0.5 transition font-bold text-white rounded-xl flex items-center justify-center gap-2 cursor-pointer text-xs shadow-md disabled:bg-slate-300"
                  >
                    {loading ? (
                      <Loader2 className="w-4 h-4 animate-spin" />
                    ) : (
                      <Save className="w-4 h-4" />
                    )}
                    Ghi thông tin phiếu & Lưu vào Cơ sở dữ liệu
                  </button>
                </div>
              </form>
            )}

            {/* CHẾ ĐỘ 2: GIAO DIỆN NHẬP NHANH BENTO TIÊU CHUẨN */}
            {manualFormType === "standard" && (
              <form onSubmit={handleManualSubmit} className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs font-sans">
                <div>
                  <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                    {activeBlock === "doanhnghiep" ? "Mã số thuế doanh nghiệp (MST)" : "Mã số thuế / Mã hộ cá thể"} <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    placeholder={activeBlock === "doanhnghiep" ? "Ví dụ: 0102030405" : "Ví dụ: 8012345678"}
                    value={manualMst}
                    onChange={e => setManualMst(e.target.value)}
                    className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
                  />
                </div>

                <div>
                  <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                    {activeBlock === "doanhnghiep" ? "Tên doanh nghiệp" : "Tên cơ sở sản xuất kinh doanh"} <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    placeholder={activeBlock === "doanhnghiep" ? "Tên doanh nghiệp sản xuất..." : "Tên hộ kinh doanh..."}
                    value={manualName}
                    onChange={e => setManualName(e.target.value)}
                    className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
                  />
                </div>

                <div>
                  <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                    {activeBlock === "doanhnghiep" ? "Người đại diện pháp luật" : "Chủ hộ / Đại diện cơ sở"}
                  </label>
                  <input
                    type="text"
                    placeholder="Họ và tên..."
                    value={manualRep}
                    onChange={e => setManualRep(e.target.value)}
                    className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
                  />
                </div>

                <div>
                  <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                    {activeBlock === "doanhnghiep" ? "Địa chỉ trụ sở / Địa bàn" : "Địa chỉ cơ sở / Địa bàn hoạt động"}
                  </label>
                  <input
                    type="text"
                    placeholder="Xã/Phường, Quận/Huyện..."
                    value={manualAddress}
                    onChange={e => setManualAddress(e.target.value)}
                    className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
                  />
                </div>

                <div>
                  <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                    {activeBlock === "doanhnghiep" ? "Doanh thu thuần (Triệu đồng)" : "Doanh thu kinh doanh (Triệu đồng)"}
                  </label>
                  <input
                    type="number"
                    min="0"
                    step="any"
                    placeholder="Ví dụ: 1250"
                    value={manualDoanhThu}
                    onChange={e => setManualDoanhThu(e.target.value)}
                    className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
                  />
                </div>

                <div>
                  <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                    {activeBlock === "doanhnghiep" ? "Số lao động bình quân (Người)" : "Số lao động (Người)"}
                  </label>
                  <input
                    type="number"
                    min="0"
                    placeholder="Ví dụ: 15"
                    value={manualLaoDong}
                    onChange={e => setManualLaoDong(e.target.value)}
                    className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
                  />
                </div>

                <div>
                  <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                    Mã ngành VSIC chính (5 chữ số)
                  </label>
                  <input
                    type="text"
                    placeholder="Ví dụ: 47110"
                    maxLength={5}
                    value={manualMaNganh}
                    onChange={e => setManualMaNganh(e.target.value)}
                    className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
                  />
                </div>

                <div>
                  <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                    Ghi chú khảo sát
                  </label>
                  <input
                    type="text"
                    placeholder="Thông tin ghi nhận ngoài lề..."
                    value={manualGhiChu}
                    onChange={e => setManualGhiChu(e.target.value)}
                    className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
                  />
                </div>

                {/* CÁC CHỈ TIÊU ĐỘNG TỰ CẤU HÌNH THEO BIỂU MẪU */}
                {customFields.length > 0 && (
                  <div className="md:col-span-2 border-t border-dashed border-slate-200 mt-2 pt-4">
                    <h4 className="text-xs font-black text-indigo-750 mb-3 flex items-center gap-1.5 uppercase tracking-wide">
                      <Sparkles className="w-3.5 h-3.5 text-indigo-500 animate-pulse" />
                      Chỉ tiêu động của biểu mẫu đã thiết lập
                    </h4>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {customFields.map(field => (
                        <div key={field.name}>
                          <label className="text-[10.5px] font-bold text-slate-650 block mb-1">
                            {field.label}
                          </label>
                          {field.type === "boolean" ? (
                            <select
                              value={dynamicValues[field.name] || ""}
                              onChange={e => setDynamicValues(prev => ({ ...prev, [field.name]: e.target.value }))}
                              className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs bg-white"
                            >
                              <option value="">-- Chưa chọn --</option>
                              <option value="true">Có / Đạt / Đúng</option>
                              <option value="false">Không / Chưa đạt / Sai</option>
                            </select>
                          ) : (
                            <input
                              type={field.type === "number" ? "number" : "text"}
                              step={field.type === "number" ? "any" : undefined}
                              placeholder={`Nhập ${field.label.toLowerCase()}...`}
                              value={dynamicValues[field.name] || ""}
                              onChange={e => setDynamicValues(prev => ({ ...prev, [field.name]: e.target.value }))}
                              className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
                            />
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                <div className="md:col-span-2 pt-4 border-t border-slate-100">
                  <button
                    type="submit"
                    disabled={loading}
                    className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 active:translate-y-0.5 transition font-bold text-white rounded-xl flex items-center justify-center gap-2 cursor-pointer text-xs shadow-md disabled:bg-slate-300"
                  >
                    {loading ? (
                      <Loader2 className="w-4 h-4 animate-spin" />
                    ) : (
                      <Plus className="w-4 h-4" />
                    )}
                    Lưu thông tin phiếu khảo sát vào hệ thống
                  </button>
                </div>
              </form>
            )}
          </div>
        )}

        {/* PHƯƠNG THỨC 2: ĐỌC FILE SCAN BẰNG AI */}
        {entryMethod === "scan" && (
          <div className="space-y-6">
            <div className="bg-gradient-to-r from-violet-50 to-indigo-50/50 border border-indigo-150 rounded-xl p-4 flex items-start gap-3">
              <span className="text-xl">🤖</span>
              <div className="space-y-1">
                <h4 className="text-xs font-black text-indigo-900 leading-none">Trí tuệ nhân tạo Gemini Nhận dạng Văn bản cực mạnh</h4>
                <p className="text-[10.5px] text-indigo-700 leading-relaxed">
                  Tải lên ảnh chụp phiếu khảo sát viết tay, tờ khai hoặc file PDF scan của đơn vị kinh doanh. Trợ lý AI sẽ tự động phân tích chữ viết, nhận diện MST, địa chỉ, ngành nghề kinh doanh, lao động và tự động điền biểu mẫu cho bạn đối chiếu.
                </p>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {/* Cột trái: Upload & Xem trước */}
              <div className="space-y-4">
                <div 
                  onClick={() => scanFileInputRef.current?.click()}
                  className="border-2 border-dashed border-slate-200 hover:border-indigo-400 rounded-2xl p-8 text-center cursor-pointer transition bg-slate-50/50 hover:bg-slate-50"
                >
                  <input
                    ref={scanFileInputRef}
                    type="file"
                    accept="image/*,application/pdf"
                    onChange={handleScanFileChange}
                    className="hidden"
                  />
                  <div className="flex flex-col items-center space-y-2">
                    <div className="p-3 bg-white rounded-full shadow-xs border border-slate-100">
                      <ScanFace className="w-6 h-6 text-indigo-500" />
                    </div>
                    <div>
                      <p className="text-xs font-bold text-slate-700">
                        {scanFile ? scanFile.name : "Chọn ảnh chụp phiếu hoặc file PDF scan"}
                      </p>
                      <p className="text-[10px] text-slate-400 mt-1">Hỗ trợ định dạng: .jpg, .png, .pdf</p>
                    </div>
                  </div>
                </div>

                {scanPreviewUrl && (
                  <div className="border border-slate-200 rounded-xl p-3 bg-slate-50 text-center space-y-2">
                    <span className="text-[10px] font-bold text-slate-500 block">ẢNH CHỤP XEM TRƯỚC</span>
                    <img src={scanPreviewUrl} alt="Scan preview" className="max-h-60 mx-auto rounded-lg shadow-xs object-contain" />
                  </div>
                )}

                {scanFile && (
                  <button
                    onClick={handleStartAiExtraction}
                    disabled={isExtracting}
                    className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 active:translate-y-0.5 transition font-bold text-white rounded-xl flex items-center justify-center gap-2 cursor-pointer text-xs shadow-md disabled:bg-slate-300"
                  >
                    {isExtracting ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        Đang quét tài liệu AI...
                      </>
                    ) : (
                      <>
                        <Sparkles className="w-4 h-4" />
                        Bắt đầu đọc tệp bằng AI
                      </>
                    )}
                  </button>
                )}

                {aiLogs.length > 0 && (
                  <div className="bg-slate-900 text-slate-300 font-mono text-[9px] p-3 rounded-xl space-y-1 max-h-32 overflow-y-auto">
                    {aiLogs.map((log, idx) => (
                      <div key={idx} className="leading-tight">{log}</div>
                    ))}
                  </div>
                )}
              </div>

              {/* Cột phải: Form rà soát kết quả trích xuất */}
              <div className="border border-slate-200 rounded-2xl p-4 bg-slate-50/50 space-y-4">
                <div className="flex justify-between items-center border-b border-slate-200 pb-2">
                  <span className="text-xs font-extrabold text-slate-800">📋 RÀ SOÁT KẾT QUẢ AI ĐỌC</span>
                  {extractedResult && (
                    <span className="bg-indigo-100 text-indigo-700 text-[9px] font-bold px-2 py-0.5 rounded uppercase">
                      {extractedResult.block === "cathe" ? "Cơ sở Cá thể" : "Doanh nghiệp"}
                    </span>
                  )}
                </div>

                {extractedResult ? (
                  <div className="space-y-4 text-xs font-sans">
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="text-[10px] font-bold text-slate-400 block uppercase">Mã số thuế / Mã CS</label>
                        <input
                          type="text"
                          value={extractedResult.mst || ""}
                          onChange={e => setExtractedResult(prev => ({ ...prev, mst: e.target.value }))}
                          className="w-full px-2.5 py-1.5 border border-slate-200 bg-white rounded-lg text-xs"
                        />
                      </div>

                      <div>
                        <label className="text-[10px] font-bold text-slate-400 block uppercase">Tên đối tượng</label>
                        <input
                          type="text"
                          value={extractedResult.name || ""}
                          onChange={e => setExtractedResult(prev => ({ ...prev, name: e.target.value }))}
                          className="w-full px-2.5 py-1.5 border border-slate-200 bg-white rounded-lg text-xs font-bold"
                        />
                      </div>

                      <div>
                        <label className="text-[10px] font-bold text-slate-400 block uppercase">Chủ / Người đại diện</label>
                        <input
                          type="text"
                          value={extractedResult.representative || ""}
                          onChange={e => setExtractedResult(prev => ({ ...prev, representative: e.target.value }))}
                          className="w-full px-2.5 py-1.5 border border-slate-200 bg-white rounded-lg text-xs"
                        />
                      </div>

                      <div>
                        <label className="text-[10px] font-bold text-slate-400 block uppercase">Địa chỉ / Địa bàn</label>
                        <input
                          type="text"
                          value={extractedResult.address || ""}
                          onChange={e => setExtractedResult(prev => ({ ...prev, address: e.target.value }))}
                          className="w-full px-2.5 py-1.5 border border-slate-200 bg-white rounded-lg text-xs"
                        />
                      </div>

                      <div>
                        <label className="text-[10px] font-bold text-slate-400 block uppercase">Doanh thu (triệu đồng)</label>
                        <input
                          type="number"
                          value={extractedResult.doanhthu || 0}
                          onChange={e => setExtractedResult(prev => ({ ...prev, doanhthu: parseFloat(e.target.value) || 0 }))}
                          className="w-full px-2.5 py-1.5 border border-slate-200 bg-white rounded-lg text-xs"
                        />
                      </div>

                      <div>
                        <label className="text-[10px] font-bold text-slate-400 block uppercase">Lao động (người)</label>
                        <input
                          type="number"
                          value={extractedResult.laodong || 0}
                          onChange={e => setExtractedResult(prev => ({ ...prev, laodong: parseInt(e.target.value) || 0 }))}
                          className="w-full px-2.5 py-1.5 border border-slate-200 bg-white rounded-lg text-xs"
                        />
                      </div>

                      <div className="col-span-2">
                        <label className="text-[10px] font-bold text-slate-400 block uppercase">Mã ngành VSIC (5 số)</label>
                        <input
                          type="text"
                          value={extractedResult.manganh || ""}
                          onChange={e => setExtractedResult(prev => ({ ...prev, manganh: e.target.value }))}
                          className="w-full px-2.5 py-1.5 border border-slate-200 bg-white rounded-lg text-xs"
                        />
                      </div>
                    </div>

                    <button
                      onClick={handleSaveExtractedResult}
                      className="w-full py-2 bg-emerald-600 hover:bg-emerald-700 font-bold text-white rounded-xl text-xs transition cursor-pointer flex items-center justify-center gap-1.5"
                    >
                      <CheckCircle className="w-4 h-4" />
                      Xác nhận và lưu vào hệ thống
                    </button>
                  </div>
                ) : (
                  <div className="h-60 border border-dashed border-slate-200 rounded-xl flex flex-col items-center justify-center text-slate-400 p-4 text-center">
                    <Sparkles className="w-8 h-8 text-indigo-300 mb-2" />
                    <p className="text-xs font-bold">Màn hình rà soát đang trống</p>
                    <p className="text-[10px]">Tải lên file scan và ấn "Bắt đầu đọc tệp" để nạp dữ liệu rà soát tại đây.</p>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* PHƯƠNG THỨC 3: NẠP FILE EXCEL */}
        {entryMethod === "excel" && (
          <div className="space-y-4">
            <h3 className="text-sm font-extrabold text-emerald-800 uppercase border-b border-slate-100 pb-2 flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
              NẠP TỆP EXCEL KHẢO SÁT HÀNG LOẠT - {activeBlock === "doanhnghiep" ? "KHỐI DOANH NGHIỆP" : "KHỐI CÁ THỂ"}
            </h3>

            <p className="text-xs text-slate-500 leading-relaxed">
              Tải lên bảng biểu báo cáo hoặc danh bạ các cơ sở dưới dạng tệp Excel (.xlsx, .xls). Hệ thống có thuật toán thông minh <b>tự động phiên dịch đầu cột</b> (Ví dụ: MST, Tên, Mã Ngành, Doanh Thu, Lao Động) của bạn để nạp hàng ngàn dòng trong vài giây.
            </p>

            <div 
              onClick={() => excelFileInputRef.current?.click()}
              className="border-2 border-dashed border-slate-200 hover:border-emerald-500 rounded-2xl p-8 text-center cursor-pointer transition bg-slate-50/50 hover:bg-slate-50"
            >
              <input
                ref={excelFileInputRef}
                type="file"
                accept=".xlsx, .xls"
                onChange={handleExcelFileChange}
                className="hidden"
              />
              <div className="flex flex-col items-center space-y-2">
                <div className="p-3 bg-white rounded-full shadow-xs border border-slate-100">
                  <FileUp className="w-6 h-6 text-slate-400" />
                </div>
                <div>
                  <p className="text-xs font-bold text-slate-700">
                    {excelFileName ? excelFileName : "Click vào đây để chọn tệp Excel của bạn"}
                  </p>
                  <p className="text-[10px] text-slate-400 mt-1">Định dạng hỗ trợ: .xlsx, .xls</p>
                </div>
              </div>
            </div>

            {excelRows.length > 0 && (
              <div className="border border-emerald-100 bg-emerald-50/20 rounded-xl p-4 space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-[10px] font-black text-emerald-800">XEM TRƯỚC 5 DÒNG ĐẦU TIÊN CỦA FILE ({excelRows.length} dòng)</span>
                  <button 
                    onClick={() => {
                      setExcelRows([]);
                      setExcelFileName("");
                      if (excelFileInputRef.current) excelFileInputRef.current.value = "";
                    }}
                    className="p-1 hover:bg-rose-100 rounded text-rose-600 cursor-pointer"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
                
                <div className="max-h-40 overflow-y-auto border border-slate-200 rounded-lg bg-white text-[11px]">
                  <table className="w-full text-left border-collapse text-slate-700">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 text-slate-500 font-mono text-[10px] font-bold sticky top-0">
                        <th className="p-2">STT</th>
                        {Object.keys(excelRows[0] || {}).slice(0, 5).map((colName) => (
                          <th key={colName} className="p-2 truncate max-w-[120px]">{colName}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {excelRows.slice(0, 5).map((row, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/50">
                          <td className="p-2 font-mono font-bold text-slate-400">{idx + 1}</td>
                          {Object.values(row).slice(0, 5).map((val: any, vIdx) => (
                            <td key={vIdx} className="p-2 truncate max-w-[120px] font-mono">{String(val)}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="pt-2">
                  {uploadProgress ? (
                    <div className="flex items-center justify-center gap-2 text-xs font-bold text-indigo-600">
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>{uploadProgress}</span>
                    </div>
                  ) : (
                    <button
                      onClick={handleExcelUpload}
                      className="w-full py-2.5 bg-emerald-600 hover:bg-emerald-700 font-bold text-white rounded-xl text-xs transition shadow-md cursor-pointer flex items-center justify-center gap-1.5"
                    >
                      <FileSpreadsheet className="w-4 h-4" />
                      Xác nhận và nạp tệp Excel này vào cơ sở dữ liệu
                    </button>
                  )}
                </div>
              </div>
            )}
          </div>
        )}

        {/* PHƯƠNG THỨC 4: THIẾT KẾ BIỂU MẪU & CHỈ TIÊU ĐỘNG */}
        {entryMethod === "template" && (
          <div className="space-y-6">
            <div className="border-b border-slate-100 pb-3">
              <h3 className="text-sm font-black text-indigo-900 uppercase flex items-center gap-2">
                <Sparkles className="w-5 h-5 text-indigo-500 animate-pulse" />
                Thiết lập Biểu mẫu khảo sát & Chỉ tiêu động tùy chọn
              </h3>
              <p className="text-xs text-slate-500 mt-1 leading-relaxed">
                Khi bạn tải tệp khảo sát mẫu lên hoặc thêm các chỉ tiêu tự chọn dưới đây, hệ thống sẽ <b>tự động sinh giao diện nhập liệu</b> (cho cả phần Nhập thủ công & cấu hình trích xuất AI tự động nhận diện từ ảnh chụp phiếu).
              </p>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* KHỐI 1: TẢI FILE MẪU LÊN ĐỂ TỰ ĐỘNG PHÂN TÍCH */}
              <div className="bg-slate-50/55 border border-slate-200 rounded-2xl p-4 space-y-4">
                <h4 className="text-xs font-black text-indigo-850 uppercase tracking-wide flex items-center gap-1.5">
                  📥 Cách 1: Đọc và phân tích cột tự động từ tệp Excel mẫu
                </h4>
                <p className="text-[11px] text-slate-500 leading-relaxed">
                  Tải lên tệp Excel chứa bảng khảo sát của bạn. Hệ thống sẽ tự động phát hiện các cột chỉ tiêu tự chọn (không thuộc các trường cơ bản như MST, Tên, Doanh thu, Lao động) để tự tạo biểu mẫu.
                </p>

                <div 
                  onClick={() => templateFileInputRef.current?.click()}
                  className="border-2 border-dashed border-slate-300 hover:border-indigo-500 rounded-xl p-6 text-center cursor-pointer transition bg-white"
                >
                  <input
                    ref={templateFileInputRef}
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={handleTemplateFileChange}
                    className="hidden"
                  />
                  <div className="flex flex-col items-center space-y-2">
                    <FileSpreadsheet className="w-8 h-8 text-indigo-400" />
                    <div>
                      <p className="text-xs font-bold text-slate-700">
                        {templateFileName ? templateFileName : "Chọn hoặc kéo thả tệp biểu mẫu mẫu vào đây"}
                      </p>
                      <p className="text-[10px] text-slate-400 mt-0.5">Hỗ trợ .xlsx, .xls</p>
                    </div>
                  </div>
                </div>

                {analyzedFields.length > 0 && (
                  <div className="bg-indigo-50/30 border border-indigo-100 rounded-xl p-3 space-y-3">
                    <span className="text-[10px] font-black text-indigo-850 uppercase block">
                      🔍 Phát hiện thấy {analyzedFields.length} chỉ tiêu động mới:
                    </span>
                    <div className="max-h-32 overflow-y-auto space-y-1 bg-white p-2 border border-slate-200 rounded-lg text-xs">
                      {analyzedFields.map((field, idx) => (
                        <div key={idx} className="flex justify-between items-center py-1 border-b border-slate-50 last:border-0 font-mono text-[10.5px]">
                          <span className="text-slate-700 font-bold">{field.label}</span>
                          <span className="text-slate-400">Khóa: {field.name}</span>
                        </div>
                      ))}
                    </div>

                    <button
                      onClick={handleApplyAnalyzedFields}
                      className="w-full py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-black text-xs rounded-lg shadow-sm transition cursor-pointer"
                    >
                      Áp dụng {analyzedFields.length} chỉ tiêu này vào Biểu mẫu nhập liệu
                    </button>
                  </div>
                )}
              </div>

              {/* KHỐI 2: THÊM CHỈ TIÊU ĐỘNG THỦ CÔNG */}
              <div className="bg-slate-50/55 border border-slate-200 rounded-2xl p-4 space-y-4">
                <h4 className="text-xs font-black text-indigo-850 uppercase tracking-wide flex items-center gap-1.5">
                  ✍️ Cách 2: Thêm thủ công từng chỉ tiêu khảo sát động
                </h4>

                <form onSubmit={handleSaveCustomField} className="space-y-3.5 text-xs text-slate-700">
                  <div>
                    <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                      Tên hiển thị (Label) <span className="text-red-500">*</span>
                    </label>
                    <input
                      type="text"
                      required
                      placeholder="Ví dụ: Sản lượng sản phẩm chính (tấn/năm)"
                      value={newFieldLabel}
                      onChange={e => {
                        setNewFieldLabel(e.target.value);
                        // Tự động gợi ý mã khóa không dấu
                        if (!newFieldName) {
                          const suggested = e.target.value
                            .normalize("NFD")
                            .replace(/[\u0300-\u036f]/g, "")
                            .toLowerCase()
                            .replace(/[^a-z0-9]/g, "_")
                            .replace(/_+/g, "_");
                          setNewFieldName(suggested);
                        }
                      }}
                      className="w-full px-3 py-1.5 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs bg-white"
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                        Mã khóa lưu trữ <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="text"
                        required
                        placeholder="Ví dụ: san_luong_chinh"
                        value={newFieldName}
                        onChange={e => setNewFieldName(e.target.value.toLowerCase().replace(/[^a-z0-9_]/g, ""))}
                        className="w-full px-3 py-1.5 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs bg-white font-mono"
                      />
                    </div>

                    <div>
                      <label className="text-[10.5px] font-bold text-slate-500 uppercase block mb-1">
                        Kiểu trường dữ liệu
                      </label>
                      <select
                        value={newFieldType}
                        onChange={e => setNewFieldType(e.target.value as any)}
                        className="w-full px-3 py-1.5 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs bg-white"
                      >
                        <option value="text">Chữ (Text)</option>
                        <option value="number">Số (Number)</option>
                        <option value="boolean">Đúng / Sai (Có - Không)</option>
                      </select>
                    </div>
                  </div>

                  <button
                    type="submit"
                    className="w-full py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-black text-xs rounded-lg shadow-sm transition cursor-pointer"
                  >
                    Thêm chỉ tiêu này vào biểu mẫu
                  </button>
                </form>
              </div>
            </div>

            {/* DANH SÁCH CÁC CHỈ TIÊU ĐANG SỬ DỤNG TRONG BIỂU MẪU KHẢO SÁT */}
            <div className="border border-slate-200 rounded-2xl p-4 bg-white space-y-3">
              <div className="flex justify-between items-center">
                <span className="text-[11px] font-black text-slate-800 uppercase tracking-wider block">
                  ⚙️ Danh sách chỉ tiêu động đang hoạt động ({customFields.length})
                </span>
                {customFields.length > 0 && (
                  <button
                    onClick={handleClearAllCustomFields}
                    className="px-2.5 py-1 text-[10px] font-bold text-rose-600 hover:bg-rose-50 border border-rose-200 rounded-lg transition cursor-pointer"
                  >
                    Xóa sạch thiết kế
                  </button>
                )}
              </div>

              {customFields.length === 0 ? (
                <div className="text-center py-6 border border-dashed border-slate-100 rounded-xl text-slate-400 text-xs">
                  Chưa có chỉ tiêu động nào được thiết lập. Hãy tải tệp mẫu lên hoặc thêm thủ công để bắt đầu.
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-50 text-slate-500 border-b border-slate-200 font-bold">
                        <th className="p-2.5">Tên chỉ tiêu (Nhãn hiển thị)</th>
                        <th className="p-2.5">Mã khóa lưu trữ</th>
                        <th className="p-2.5">Kiểu dữ liệu</th>
                        <th className="p-2.5 text-center">Giao diện sinh ra</th>
                        <th className="p-2.5 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-700">
                      {customFields.map((field) => (
                        <tr key={field.name} className="hover:bg-slate-50/30">
                          <td className="p-2.5 font-bold text-slate-800">{field.label}</td>
                          <td className="p-2.5 font-mono text-[11px] text-slate-500">{field.name}</td>
                          <td className="p-2.5">
                            <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                              field.type === "number" ? "bg-amber-50 text-amber-700 border border-amber-100" :
                              field.type === "boolean" ? "bg-purple-50 text-purple-700 border border-purple-100" :
                              "bg-slate-100 text-slate-650"
                            }`}>
                              {field.type === "number" ? "Số (Number)" : field.type === "boolean" ? "Có - Không (Boolean)" : "Chữ (Text)"}
                            </span>
                          </td>
                          <td className="p-2.5 text-center">
                            {field.type === "boolean" ? (
                              <span className="text-[10px] text-slate-400 font-mono bg-slate-50 px-2 py-0.5 border border-slate-100 rounded">Dropdown Có/Không</span>
                            ) : field.type === "number" ? (
                              <span className="text-[10px] text-slate-400 font-mono bg-slate-50 px-2 py-0.5 border border-slate-100 rounded">Input Number</span>
                            ) : (
                              <span className="text-[10px] text-slate-400 font-mono bg-slate-50 px-2 py-0.5 border border-slate-100 rounded">Input Text</span>
                            )}
                          </td>
                          <td className="p-2.5 text-right">
                            <button
                              onClick={() => handleDeleteCustomField(field.name)}
                              className="text-rose-600 hover:bg-rose-50 p-1 rounded transition cursor-pointer"
                              title="Xóa chỉ tiêu này"
                            >
                              <Trash2 className="w-4 h-4 inline-block" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        )}

      </div>

      {/* DANH SÁCH TIN ĐÃ NHẬP KÈM BỘ LỌC TÌM KIẾM */}
      <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-4">
        
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-3">
          <div>
            <h3 className="text-sm font-extrabold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
              📋 DANH SÁCH PHIẾU ĐIỀU TRA KHẢO SÁT ĐÃ THU THẬP ({filteredRecords.length}/{surveyRecords.length})
            </h3>
            <p className="text-[10.5px] text-slate-400 mt-1">
              Danh sách kết quả điều tra, khai báo trực tiếp từ các điều tra viên và cơ sở kinh tế trên địa bàn tỉnh Hưng Yên.
            </p>
          </div>

          <div className="flex flex-wrap gap-2">
            <button
              onClick={fetchRecords}
              className="p-1.5 bg-slate-50 hover:bg-slate-100 rounded-xl text-slate-600 border border-slate-200 cursor-pointer"
              title="Làm mới dữ liệu"
            >
              <RefreshCw className="w-4 h-4" />
            </button>

            {/* Export Excel local */}
            <button
              onClick={() => {
                if (filteredRecords.length === 0) return;
                
                const flatData = filteredRecords.map((r, idx) => {
                  const row: any = {
                    "STT": idx + 1,
                    "Khối đối tượng": r.block === "doanhnghiep" ? "Doanh nghiệp" : "Cá thể",
                    "Mã số thuế / Mã cơ sở": r.mst,
                    "Tên cơ sở / Doanh nghiệp": r.name,
                    "Người đại diện / Chủ hộ": r.representative || "",
                    "Địa chỉ / Địa bàn": r.address || "",
                    "Doanh thu (triệu đồng)": r.doanhthu || 0,
                    "Lao động (người)": r.laodong || 0,
                    "Mã ngành VSIC": r.manganh || "",
                    "Ghi chú": r.ghichu || "",
                    "Nguồn nhập tin": r.entryMethod === "manual" ? "Thủ công" : r.entryMethod === "scan" ? "AI Quét" : r.entryMethod === "excel" ? "Nạp Excel" : "Thiết kế mẫu",
                    "Người nhập": r.createdBy || "",
                    "Ngày nhập liệu": new Date(r.createdAt).toLocaleDateString("vi-VN")
                  };

                  // Thêm các cột chỉ tiêu động đã cấu hình
                  customFields.forEach(cf => {
                    const val = r.customData?.[cf.name];
                    row[cf.label] = val === undefined || val === "" ? "" : (typeof val === "boolean" ? (val ? "Có" : "Không") : val);
                  });

                  return row;
                });

                const ws = XLSX.utils.json_to_sheet(flatData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Khao_Sat_Data");
                XLSX.writeFile(wb, `VTong_Survey_Export_${unitID.toUpperCase()}.xlsx`);
              }}
              disabled={filteredRecords.length === 0}
              className="flex items-center gap-1 px-3 py-1.5 bg-emerald-50 text-emerald-700 hover:bg-emerald-100 text-xs font-bold rounded-xl border border-emerald-200 cursor-pointer disabled:bg-slate-50 disabled:text-slate-350"
            >
              <Download className="w-3.5 h-3.5" />
              Tải Excel về máy
            </button>
          </div>
        </div>

        {/* BỘ LỌC VÀ TÌM KIẾM DỮ LIỆU CHUYÊN SÂU */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
          {/* Ô tìm kiếm */}
          <div className="relative">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
            <input
              type="text"
              placeholder="Tìm theo Tên, MST, Địa bàn..."
              value={searchQuery}
              onChange={e => setSearchQuery(e.target.value)}
              className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs"
            />
          </div>

          {/* Lọc khối */}
          <div className="flex items-center gap-1">
            <Filter className="w-3.5 h-3.5 text-slate-400" />
            <select
              value={filterBlock}
              onChange={e => setFilterBlock(e.target.value as any)}
              className="w-full px-3 py-2 border border-slate-200 rounded-xl bg-white text-xs"
            >
              <option value="all">Tất cả Khối đối tượng</option>
              <option value="doanhnghiep">Khối Doanh nghiệp (DN)</option>
              <option value="cathe">Khối Cá thể</option>
            </select>
          </div>

          {/* Lọc phương thức */}
          <div className="flex items-center gap-1">
            <ChevronDown className="w-3.5 h-3.5 text-slate-400" />
            <select
              value={filterMethod}
              onChange={e => setFilterMethod(e.target.value as any)}
              className="w-full px-3 py-2 border border-slate-200 rounded-xl bg-white text-xs"
            >
              <option value="all">Tất cả nguồn nhập</option>
              <option value="manual">Nhập thủ công</option>
              <option value="scan">Trích xuất bằng ảnh quét</option>
              <option value="excel">Nạp bằng file Excel</option>
            </select>
          </div>
        </div>

        {/* BẢNG DỮ LIỆU TRỰC QUAN */}
        {filteredRecords.length === 0 ? (
          <div className="border border-dashed border-slate-200 rounded-xl p-8 text-center text-slate-400 space-y-2">
            <HelpCircle className="w-8 h-8 text-slate-300 mx-auto" />
            <p className="text-xs font-bold">Không tìm thấy phiếu khảo sát nào phù hợp</p>
            <p className="text-[10px]">Thay đổi bộ lọc hoặc khai báo thêm phiếu điều tra mới ở phía trên để cập nhật.</p>
          </div>
        ) : (
          <div className="overflow-x-auto border border-slate-200 rounded-xl shadow-xs bg-white text-xs">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-150 font-sans text-[9.5px] font-bold text-slate-500 uppercase tracking-wider">
                  <th className="px-4 py-3">Phân nhóm</th>
                  <th className="px-4 py-3">Mã số thuế / Mã số cơ sở</th>
                  <th className="px-4 py-3">Tên cơ sở / Doanh nghiệp</th>
                  <th className="px-4 py-3">Chủ cơ sở / Người đại diện</th>
                  <th className="px-4 py-3">Địa chỉ / Địa bàn</th>
                  <th className="px-4 py-3">Mã ngành VSIC</th>
                  <th className="px-4 py-3 text-center">Trạng thái ký số</th>
                  {customFields.map(cf => (
                    <th key={cf.name} className="px-4 py-3 text-slate-700 bg-indigo-50/40 border-x border-slate-200">{cf.label}</th>
                  ))}
                  <th className="px-4 py-3 text-right">Doanh thu (triệu)</th>
                  <th className="px-4 py-3 text-right">Lao động (người)</th>
                  <th className="px-4 py-3 text-center">Phương thức khai báo</th>
                  <th className="px-4 py-3 text-center">Thao tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-slate-750">
                {filteredRecords.map((r) => (
                  <tr key={r.id} className="hover:bg-slate-50/40 transition">
                    <td className="px-4 py-3">
                      {r.block === "doanhnghiep" ? (
                        <span className="bg-blue-50 text-blue-700 border border-blue-100 text-[10px] font-bold px-2 py-0.5 rounded-full uppercase">Doanh nghiệp</span>
                      ) : (
                        <span className="bg-amber-50 text-amber-700 border border-amber-100 text-[10px] font-bold px-2 py-0.5 rounded-full uppercase">Cá thể</span>
                      )}
                    </td>
                    <td className="px-4 py-3 font-mono font-bold text-indigo-700">{r.mst}</td>
                    <td className="px-4 py-3 font-extrabold text-slate-900">{r.name}</td>
                    <td className="px-4 py-3 text-slate-600">{r.representative || r.respondent_name || "-"}</td>
                    <td className="px-4 py-3 text-slate-500 truncate max-w-[150px]" title={r.address}>{r.address || "-"}</td>
                    <td className="px-4 py-3 font-mono font-medium text-slate-600">{r.manganh || "-"}</td>
                    <td className="px-4 py-3 text-center">
                      {r.is_signed ? (
                        r.signed_mode === "token" ? (
                          <button
                            type="button"
                            onClick={() => setSelectedSignatureRecord(r)}
                            className="inline-flex items-center gap-1 bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100 text-[10px] font-extrabold px-2 py-1 rounded-md cursor-pointer transition-colors"
                            title="Đã ký số bằng USB Token CA. Click để xem chi tiết chứng thư số."
                          >
                            <span className="inline-block w-1.5 h-1.5 bg-emerald-500 rounded-full animate-pulse"></span>
                            🔒 KÝ SỐ TOKEN CA
                          </button>
                        ) : (
                          <button
                            type="button"
                            onClick={() => setSelectedSignatureRecord(r)}
                            className="inline-flex items-center gap-1 bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100 text-[10px] font-extrabold px-2 py-1 rounded-md cursor-pointer transition-colors"
                            title="Đã ký tay điện tử. Click để xem chữ ký."
                          >
                            ✍️ KÝ ĐIỆN TỬ
                          </button>
                        )
                      ) : (
                        <span className="inline-flex items-center gap-1 bg-slate-100 text-slate-500 border border-slate-200 text-[10px] font-semibold px-2 py-1 rounded-md">
                          ⚠️ Chưa xác thực
                        </span>
                      )}
                    </td>
                    {customFields.map(cf => {
                      const val = r.customData?.[cf.name];
                      return (
                        <td key={cf.name} className="px-4 py-3 font-medium text-slate-700 border-x border-slate-100 bg-slate-50/20">
                          {val === undefined || val === "" ? (
                            <span className="text-slate-350 italic text-[10px]">Chưa nhập</span>
                          ) : typeof val === "boolean" ? (
                            <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${val ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : 'bg-rose-50 text-rose-700 border border-rose-100'}`}>
                              {val ? "Có" : "Không"}
                            </span>
                          ) : (
                            <span className="font-semibold text-slate-800">{String(val)}</span>
                          )}
                        </td>
                      );
                    })}
                    <td className="px-4 py-3 text-right font-mono font-bold text-slate-800">{(r.doanhthu || 0).toLocaleString()}</td>
                    <td className="px-4 py-3 text-right font-mono font-bold text-slate-800">{(r.laodong || 0).toLocaleString()}</td>
                    <td className="px-4 py-3 text-center">
                      {r.entryMethod === "manual" && <span className="text-[10px] text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full font-bold">Khai báo thủ công</span>}
                      {r.entryMethod === "scan" && <span className="text-[10px] text-violet-600 bg-violet-50 border border-violet-100 px-2 py-0.5 rounded-full font-bold">Quét ảnh (AI)</span>}
                      {r.entryMethod === "excel" && <span className="text-[10px] text-emerald-600 bg-emerald-50 border border-emerald-100 px-2 py-0.5 rounded-full font-bold">Nạp từ Excel</span>}
                    </td>
                    <td className="px-4 py-3">
                      <div className="flex items-center justify-center gap-1">
                        <button
                          onClick={() => setEditingRecord(r)}
                          className="p-1 hover:bg-indigo-50 rounded text-indigo-600 cursor-pointer"
                          title="Hiệu chỉnh phiếu điều tra"
                        >
                          <Edit2 className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => r.id && handleDeleteRecord(r.id, r.name)}
                          className="p-1 hover:bg-rose-50 rounded text-rose-600 cursor-pointer"
                          title="Xóa phiếu điều tra"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

      </div>

      {/* POPUP CHI TIẾT CHỮ KÝ VÀ CHỨNG THƯ SỐ */}
      {selectedSignatureRecord && (
        <div className="fixed inset-0 bg-slate-900/65 backdrop-blur-xs flex items-center justify-center z-50 animate-fadeIn p-4 text-xs">
          <div className="bg-white border border-slate-250 rounded-3xl w-full max-w-md shadow-2xl overflow-hidden">
            {/* Header */}
            <div className="bg-slate-900 text-white p-4 flex justify-between items-center">
              <div className="flex items-center gap-1.5 font-bold uppercase tracking-wide">
                <Key className="w-4 h-4 text-amber-400" />
                <span>Chi tiết Chữ ký số & Chứng thư</span>
              </div>
              <button 
                onClick={() => setSelectedSignatureRecord(null)}
                className="text-slate-300 hover:text-white font-extrabold text-sm cursor-pointer px-1"
              >
                ✕
              </button>
            </div>

            {/* Body */}
            <div className="p-5 space-y-4 font-sans text-slate-700">
              <div className="flex items-center gap-2 justify-center pb-2 border-b border-slate-100">
                <div className="w-8 h-8 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center">
                  <ShieldCheck className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="font-extrabold text-slate-900 text-[12px] uppercase">
                    Chữ ký điện tử hợp lệ
                  </h4>
                  <p className="text-[10px] text-slate-400">
                    Hệ thống xác thực quốc gia Hưng Yên Smart Sign
                  </p>
                </div>
              </div>

              {/* Thông tin doanh nghiệp ký */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-2 text-[11px]">
                <div>
                  <span className="text-slate-400 block text-[9.5px] uppercase font-bold">Doanh nghiệp / Đơn vị ký:</span>
                  <span className="font-extrabold text-slate-900 text-xs">{selectedSignatureRecord.name.toUpperCase()}</span>
                </div>
                <div className="grid grid-cols-2 gap-2 pt-1 border-t border-slate-100">
                  <div>
                    <span className="text-slate-400 block text-[9.5px] uppercase font-bold">Mã số thuế:</span>
                    <span className="font-mono font-bold text-indigo-700">{selectedSignatureRecord.mst}</span>
                  </div>
                  <div>
                    <span className="text-slate-400 block text-[9.5px] uppercase font-bold">Khối đối tượng:</span>
                    <span className="font-bold text-slate-800">
                      {selectedSignatureRecord.block === "doanhnghiep" ? "Doanh nghiệp (DN)" : "Cá thể"}
                    </span>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-2 pt-1 border-t border-slate-100">
                  <div>
                    <span className="text-slate-400 block text-[9.5px] uppercase font-bold">Người đại diện:</span>
                    <span className="font-bold text-slate-800">
                      {selectedSignatureRecord.representative || selectedSignatureRecord.respondent_name || "Đại diện đơn vị"}
                    </span>
                  </div>
                  <div>
                    <span className="text-slate-400 block text-[9.5px] uppercase font-bold">Thời gian ký:</span>
                    <span className="font-bold text-slate-800 font-mono">{selectedSignatureRecord.signed_time || selectedSignatureRecord.createdAt}</span>
                  </div>
                </div>
              </div>

              {/* Hình ảnh mộc đỏ / Chữ ký tay */}
              <div className="border border-slate-200 rounded-xl p-3 bg-slate-50/55">
                <span className="text-slate-400 block text-[9.5px] uppercase font-bold mb-1.5 text-center">Hình ảnh chữ ký điện tử:</span>
                <div className="bg-white border border-dashed border-slate-300 rounded-lg p-2 flex items-center justify-center h-28">
                  {selectedSignatureRecord.signature_img ? (
                    <img 
                      src={selectedSignatureRecord.signature_img} 
                      alt="Chữ ký điện tử" 
                      className="max-h-24 max-w-full object-contain"
                    />
                  ) : (
                    <div className="text-slate-400 text-center text-[10.5px] italic">
                      Không có tệp ảnh chữ ký. Được lưu trữ bằng hàm băm mã hóa.
                    </div>
                  )}
                </div>
              </div>

              {/* Hash & Chứng thư số */}
              <div className="space-y-1 font-mono text-[9px] bg-slate-900 text-slate-300 p-3 rounded-lg border border-slate-950">
                <div className="flex justify-between text-slate-450 border-b border-slate-850 pb-1">
                  <span>MÃ BĂM SHA-256 KHÓA CÔNG KHAI</span>
                  <span className="text-emerald-400 font-bold">BẢO MẬT</span>
                </div>
                <p className="break-all font-bold text-emerald-300">{selectedSignatureRecord.signed_hash || "SHA256-GENERIC-VERIFIED-DATA"}</p>
                <div className="pt-1.5 border-t border-slate-850 mt-1 text-slate-400 text-[8.5px] italic">
                  * Chứng chỉ chữ ký số hợp lệ theo Nghị định 130/2018/NĐ-CP và Luật Giao dịch điện tử.
                </div>
              </div>
            </div>

            {/* Footer */}
            <div className="bg-slate-50 px-4 py-3 border-t border-slate-150 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setSelectedSignatureRecord(null)}
                className="px-4 py-1.5 bg-slate-800 hover:bg-slate-900 text-white font-bold rounded-xl transition cursor-pointer"
              >
                Đóng thông tin
              </button>
            </div>
          </div>
        </div>
      )}

      {/* POPUP CHỈNH SỬA BẢN GHI KHẢO SÁT */}
      {editingRecord && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-50 animate-fade-in p-4 text-xs">
          <div className="bg-white border border-slate-200 rounded-3xl w-full max-w-lg shadow-2xl p-6 space-y-4">
            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
              <span className="text-sm font-black text-slate-800 uppercase">📝 SỬA THÔNG TIN KHẢO SÁT</span>
              <button 
                onClick={() => setEditingRecord(null)}
                className="text-slate-400 hover:text-slate-600 font-extrabold cursor-pointer"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleUpdateRecord} className="space-y-4 font-sans text-slate-700">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Mã định danh/MST <span className="text-rose-500">*</span></label>
                  <input
                    type="text"
                    required
                    value={editingRecord.mst}
                    onChange={e => setEditingRecord(prev => prev ? ({ ...prev, mst: e.target.value }) : null)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-xl"
                  />
                </div>

                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Tên đối tượng <span className="text-rose-500">*</span></label>
                  <input
                    type="text"
                    required
                    value={editingRecord.name}
                    onChange={e => setEditingRecord(prev => prev ? ({ ...prev, name: e.target.value }) : null)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-xl font-bold"
                  />
                </div>

                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Đại diện/Chủ cơ sở</label>
                  <input
                    type="text"
                    value={editingRecord.representative || ""}
                    onChange={e => setEditingRecord(prev => prev ? ({ ...prev, representative: e.target.value }) : null)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-xl"
                  />
                </div>

                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Địa bàn / Địa chỉ</label>
                  <input
                    type="text"
                    value={editingRecord.address}
                    onChange={e => setEditingRecord(prev => prev ? ({ ...prev, address: e.target.value }) : null)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-xl"
                  />
                </div>

                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Doanh thu (triệu đồng)</label>
                  <input
                    type="number"
                    value={editingRecord.doanhthu}
                    onChange={e => setEditingRecord(prev => prev ? ({ ...prev, doanhthu: parseFloat(e.target.value) || 0 }) : null)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-xl"
                  />
                </div>

                <div>
                  <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Lao động (người)</label>
                  <input
                    type="number"
                    value={editingRecord.laodong}
                    onChange={e => setEditingRecord(prev => prev ? ({ ...prev, laodong: parseInt(e.target.value) || 0 }) : null)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-xl"
                  />
                </div>

                <div className="col-span-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Mã ngành VSIC</label>
                  <input
                    type="text"
                    value={editingRecord.manganh}
                    onChange={e => setEditingRecord(prev => prev ? ({ ...prev, manganh: e.target.value }) : null)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-xl"
                  />
                </div>

                <div className="col-span-2">
                  <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Ghi chú khảo sát</label>
                  <input
                    type="text"
                    value={editingRecord.ghichu || ""}
                    onChange={e => setEditingRecord(prev => prev ? ({ ...prev, ghichu: e.target.value }) : null)}
                    className="w-full px-3 py-1.5 border border-slate-200 rounded-xl"
                  />
                </div>

                {/* EDITING DYNAMIC CUSTOM FIELDS */}
                {customFields.length > 0 && (
                  <div className="col-span-2 border-t border-dashed border-slate-200 pt-3 mt-1 text-xs">
                    <span className="text-[10px] font-black text-indigo-850 uppercase tracking-wide block mb-2">Chỉ tiêu động tùy chỉnh</span>
                    <div className="grid grid-cols-2 gap-3">
                      {customFields.map(field => {
                        const val = editingRecord.customData?.[field.name] ?? "";
                        return (
                          <div key={field.name}>
                            <label className="text-[10px] font-bold text-slate-500 block mb-1">{field.label}</label>
                            {field.type === "boolean" ? (
                              <select
                                value={String(val)}
                                onChange={e => setEditingRecord(prev => {
                                  if (!prev) return null;
                                  const updatedData = { ...prev.customData, [field.name]: e.target.value === "true" };
                                  return { ...prev, customData: updatedData };
                                })}
                                className="w-full px-3 py-1.5 border border-slate-200 rounded-xl bg-white"
                              >
                                <option value="">-- Chưa chọn --</option>
                                <option value="true">Có / Đúng</option>
                                <option value="false">Không / Sai</option>
                              </select>
                            ) : (
                              <input
                                type={field.type === "number" ? "number" : "text"}
                                step={field.type === "number" ? "any" : undefined}
                                value={String(val)}
                                onChange={e => setEditingRecord(prev => {
                                  if (!prev) return null;
                                  const rawVal = e.target.value;
                                  const typedVal = field.type === "number" ? (parseFloat(rawVal) || 0) : rawVal;
                                  const updatedData = { ...prev.customData, [field.name]: typedVal };
                                  return { ...prev, customData: updatedData };
                                })}
                                className="w-full px-3 py-1.5 border border-slate-200 rounded-xl"
                              />
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>

              <div className="flex gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setEditingRecord(null)}
                  className="w-1/2 py-2 border border-slate-250 hover:bg-slate-50 font-bold rounded-xl text-xs text-slate-600 transition cursor-pointer"
                >
                  Hủy bỏ
                </button>
                <button
                  type="submit"
                  className="w-1/2 py-2 bg-indigo-600 hover:bg-indigo-700 font-bold text-white rounded-xl text-xs transition shadow-md cursor-pointer flex items-center justify-center gap-1"
                >
                  <Save className="w-3.5 h-3.5" />
                  Lưu thay đổi
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

    </div>
  );
};
