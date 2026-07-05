import React, { useState, useRef } from "react";
import * as XLSX from "xlsx";
import { useAuth, db, isFirebaseInitialized } from "../context/AuthContext";
import { 
  collection, 
  addDoc, 
  serverTimestamp, 
  writeBatch, 
  doc 
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
  ListFilter
} from "lucide-react";

export const DataEntry: React.FC = () => {
  const { user } = useAuth();
  const unitID = user?.unitID || "guest";
  const userName = user?.displayName || "Người dùng Khách";

  // State nhập liệu thủ công (Manual Input State)
  const [itemName, setItemName] = useState("");
  const [itemCategory, setItemCategory] = useState("Vật tư");
  const [itemValue, setItemValue] = useState("");
  const [itemQuantity, setItemQuantity] = useState("");
  const [itemNote, setItemNote] = useState("");

  // State nạp tệp Excel (Excel State)
  const [excelRows, setExcelRows] = useState<any[]>([]);
  const [excelFileName, setExcelFileName] = useState("");
  const [uploadProgress, setUploadProgress] = useState<string | null>(null);

  // Trạng thái hệ thống (System Status State)
  const [loading, setLoading] = useState(false);
  const [statusMessage, setStatusMessage] = useState<{ type: "success" | "error"; text: string } | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // ==========================================
  // LƯU NHẬT KÝ HOẠT ĐỘNG (SAVE ACTIVITY LOGS)
  // ==========================================
  // Hàm ghi nhật ký vào Firestore để AdminDashboard thu thập thông tin thời gian thực
  const logActivity = async (actionType: "manual_entry" | "excel_import", details: string) => {
    const timestamp = new Date().toISOString();
    const logData = {
      unitID,
      userName,
      actionType,
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
      // Ghi nhật ký vào LocalStorage ở chế độ mô phỏng
      const existingLogs = JSON.parse(localStorage.getItem("system_mock_logs") || "[]");
      existingLogs.unshift(logData);
      localStorage.setItem("system_mock_logs", JSON.stringify(existingLogs));
      // Gửi event để AdminDashboard đồng bộ ngay lập tức nếu đang mở song song
      window.dispatchEvent(new Event("storage_logs_updated"));
    }
  };

  // ==========================================
  // XỬ LÝ NHẬP LIỆU THỦ CÔNG (MANUAL INSERT)
  // ==========================================
  const handleManualSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!itemName.trim() || !itemValue.trim() || !itemQuantity.trim()) {
      setStatusMessage({ type: "error", text: "Vui lòng nhập đầy đủ các thông tin bắt buộc!" });
      return;
    }

    setLoading(true);
    setStatusMessage(null);

    const dataPayload = {
      name: itemName.trim(),
      category: itemCategory,
      value: parseFloat(itemValue) || 0,
      quantity: parseInt(itemQuantity) || 0,
      note: itemNote.trim(),
      createdBy: userName,
      createdAt: isFirebaseInitialized ? serverTimestamp() : new Date().toISOString()
    };

    try {
      if (isFirebaseInitialized && db) {
        // Ghi trực tiếp vào Firestore: /data/unitID/records
        // Path thực tế: /data/{unitID}/records/{documentID}
        const collectionRef = collection(db, "data", unitID, "records");
        await addDoc(collectionRef, dataPayload);
      } else {
        // Ghi vào LocalStorage ở chế độ mô phỏng
        const mockDbKey = `mock_data_${unitID}`;
        const existingData = JSON.parse(localStorage.getItem(mockDbKey) || "[]");
        existingData.unshift({ id: "item_" + Date.now(), ...dataPayload });
        localStorage.setItem(mockDbKey, JSON.stringify(existingData));
      }

      await logActivity("manual_entry", `Đã nhập thủ công một bản ghi vật tư: "${itemName}"`);

      setStatusMessage({ type: "success", text: "Lưu dữ liệu thành công!" });
      // Reset form
      setItemName("");
      setItemValue("");
      setItemQuantity("");
      setItemNote("");
    } catch (err: any) {
      console.error("Lỗi khi lưu dữ liệu:", err);
      setStatusMessage({ type: "error", text: `Lỗi khi lưu dữ liệu: ${err.message || err}` });
    } finally {
      setLoading(false);
    }
  };

  // ==========================================
  // ĐỌC VÀ XỬ LÝ FILE EXCEL (XLSX READER)
  // ==========================================
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
        
        // Chuyển đổi bảng tính sang JSON dạng Array
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        setExcelRows(jsonData);
        setStatusMessage({ type: "success", text: `Đọc tệp Excel thành công! Đã phát hiện ${jsonData.length} dòng.` });
      } catch (err: any) {
        console.error("Lỗi đọc tệp Excel:", err);
        setStatusMessage({ type: "error", text: "Tệp Excel không đúng định dạng hoặc bị hỏng." });
      }
    };
    reader.readAsBinaryString(file);
  };

  // ==========================================
  // ĐẨY EXCEL LÊN CLOUD (UPLOAD TO FIREBASE)
  // ==========================================
  const handleExcelUpload = async () => {
    if (excelRows.length === 0) {
      setStatusMessage({ type: "error", text: "Chưa có dữ liệu Excel để tải lên!" });
      return;
    }

    setLoading(true);
    setUploadProgress("Đang khởi động tải lên...");
    setStatusMessage(null);

    try {
      if (isFirebaseInitialized && db) {
        // Chia lô dữ liệu (Batch Writes) để đạt hiệu năng tải lên tốt nhất (Firebase giới hạn tối đa 500 bản ghi mỗi Batch)
        const batchSize = 100;
        const totalRows = excelRows.length;
        
        for (let i = 0; i < totalRows; i += batchSize) {
          const batch = writeBatch(db);
          const chunk = excelRows.slice(i, i + batchSize);

          chunk.forEach((row, index) => {
            const docRef = doc(collection(db, "data", unitID, "records"));
            batch.set(docRef, {
              ...row,
              uploadedBy: userName,
              uploadedAt: serverTimestamp()
            });
          });

          setUploadProgress(`Đang lưu lô ${Math.floor(i / batchSize) + 1} (${Math.min(i + batchSize, totalRows)}/${totalRows} dòng)...`);
          await batch.commit();
        }
      } else {
        // Mô phỏng lưu hàng loạt
        const mockDbKey = `mock_data_${unitID}`;
        const existingData = JSON.parse(localStorage.getItem(mockDbKey) || "[]");
        const formattedRows = excelRows.map((r, idx) => ({
          id: `excel_${Date.now()}_${idx}`,
          ...r,
          uploadedBy: userName,
          uploadedAt: new Date().toISOString()
        }));
        localStorage.setItem(mockDbKey, JSON.stringify([...formattedRows, ...existingData]));
      }

      await logActivity("excel_import", `Đã nhập khẩu tệp Excel: "${excelFileName}" với tổng cộng ${excelRows.length} dòng dữ liệu.`);
      
      setStatusMessage({ type: "success", text: `Đã nạp thành công toàn bộ ${excelRows.length} dòng dữ liệu vào kho lưu trữ của đơn vị!` });
      setExcelRows([]);
      setExcelFileName("");
      if (fileInputRef.current) fileInputRef.current.value = "";
    } catch (err: any) {
      console.error("Lỗi khi ghi dữ liệu Excel lên Cloud:", err);
      setStatusMessage({ type: "error", text: `Gặp sự cố khi đồng bộ: ${err.message || err}` });
    } finally {
      setLoading(false);
      setUploadProgress(null);
    }
  };

  return (
    <div className="space-y-8 font-sans">
      
      {/* Banner thông báo trạng thái kết nối Cloud */}
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between p-4 bg-slate-50 border border-slate-200 rounded-xl gap-3">
        <div className="flex items-center gap-2">
          <Database className={`w-5 h-5 ${isFirebaseInitialized ? "text-emerald-500" : "text-slate-400"}`} />
          <div>
            <span className="text-xs text-slate-500 block">Đơn vị đang thao tác</span>
            <span className="text-sm font-bold text-slate-800 font-mono">{unitID.toUpperCase()} ({userName})</span>
          </div>
        </div>

        <div>
          {isFirebaseInitialized ? (
            <span className="px-3 py-1 bg-emerald-50 text-emerald-600 border border-emerald-200 text-xs font-semibold rounded-full flex items-center gap-1">
              <CheckCircle className="w-3.5 h-3.5" />
              Đã đồng bộ Cloud Firebase (Real-time)
            </span>
          ) : (
            <span className="px-3 py-1 bg-amber-50 text-amber-600 border border-amber-200 text-xs font-semibold rounded-full flex items-center gap-1">
              <AlertCircle className="w-3.5 h-3.5" />
              Lưu trữ Offline (LocalStorage)
            </span>
          )}
        </div>
      </div>

      {statusMessage && (
        <div className={`p-4 rounded-xl flex items-start gap-3 border ${
          statusMessage.type === "success" 
            ? "bg-emerald-50 border-emerald-200 text-emerald-800" 
            : "bg-rose-50 border-rose-200 text-rose-800"
        }`}>
          {statusMessage.type === "success" ? (
            <CheckCircle className="w-5 h-5 shrink-0 mt-0.5 text-emerald-600" />
          ) : (
            <AlertCircle className="w-5 h-5 shrink-0 mt-0.5 text-rose-600" />
          )}
          <span className="text-sm font-medium">{statusMessage.text}</span>
        </div>
      )}

      {/* Grid chia 2 khu vực: Nhập thủ công và Đọc Excel */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        
        {/* KHU VỰC 1: FORM NHẬP THỦ CÔNG */}
        <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-4">
          <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2 border-b border-slate-100 pb-3">
            <Plus className="w-5 h-5 text-indigo-600" />
            Khai báo bản ghi thủ công
          </h3>

          <form onSubmit={handleManualSubmit} className="space-y-4 text-slate-700">
            <div>
              <label className="text-xs font-bold text-slate-500 uppercase block mb-1">Tên tài sản/Vật tư <span className="text-rose-500">*</span></label>
              <input
                type="text"
                required
                placeholder="Nhập tên vật tư..."
                value={itemName}
                onChange={e => setItemName(e.target.value)}
                className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
              />
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="text-xs font-bold text-slate-500 uppercase block mb-1">Phân loại</label>
                <select
                  value={itemCategory}
                  onChange={e => setItemCategory(e.target.value)}
                  className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm bg-white"
                >
                  <option value="Vật tư">Vật tư</option>
                  <option value="Thiết bị">Thiết bị</option>
                  <option value="Nhân lực">Nhân lực</option>
                  <option value="Dịch vụ">Dịch vụ</option>
                </select>
              </div>

              <div>
                <label className="text-xs font-bold text-slate-500 uppercase block mb-1">Số lượng <span className="text-rose-500">*</span></label>
                <input
                  type="number"
                  required
                  min="1"
                  placeholder="10"
                  value={itemQuantity}
                  onChange={e => setItemQuantity(e.target.value)}
                  className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
                />
              </div>
            </div>

            <div>
              <label className="text-xs font-bold text-slate-500 uppercase block mb-1">Đơn giá trị (VNĐ) <span className="text-rose-500">*</span></label>
              <input
                type="number"
                required
                min="0"
                placeholder="500000"
                value={itemValue}
                onChange={e => setItemValue(e.target.value)}
                className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
              />
            </div>

            <div>
              <label className="text-xs font-bold text-slate-500 uppercase block mb-1">Ghi chú bổ sung</label>
              <textarea
                rows={3}
                placeholder="Nội dung thuyết minh..."
                value={itemNote}
                onChange={e => setItemNote(e.target.value)}
                className="w-full px-3.5 py-2 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
              />
            </div>

            <button
              type="submit"
              disabled={loading}
              className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-500 active:translate-y-0.5 transition font-bold text-white rounded-xl flex items-center justify-center gap-2 cursor-pointer text-sm shadow-md disabled:bg-slate-300"
            >
              {loading ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <Plus className="w-4 h-4" />
              )}
              Lưu bản ghi vào hệ thống
            </button>
          </form>
        </div>

        {/* KHU VỰC 2: ĐỌC VÀ NHẬP KHẨU FILE EXCEL */}
        <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 flex flex-col justify-between">
          <div className="space-y-4">
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2 border-b border-slate-100 pb-3">
              <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
              Nhập tệp dữ liệu hàng loạt (Excel)
            </h3>

            <p className="text-xs text-slate-500 leading-relaxed">
              Tải lên bảng tính tài sản hoặc báo cáo định kỳ dạng Excel (.xlsx, .xls). Hệ thống sẽ tự động quét và phân tách dữ liệu để đẩy lên thư mục riêng biệt của đơn vị tại Cloud Firestore.
            </p>

            {/* Dropzone Upload file */}
            <div 
              onClick={() => fileInputRef.current?.click()}
              className="border-2 border-dashed border-slate-200 hover:border-indigo-500 rounded-2xl p-8 text-center cursor-pointer transition bg-slate-50/50 hover:bg-slate-50"
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx, .xls"
                onChange={handleExcelFileChange}
                className="hidden"
              />
              <div className="flex flex-col items-center space-y-2">
                <div className="p-3 bg-white rounded-full shadow-sm border border-slate-100">
                  <FileUp className="w-6 h-6 text-slate-400" />
                </div>
                <div>
                  <p className="text-sm font-bold text-slate-700">
                    {excelFileName ? excelFileName : "Click để chọn File Excel"}
                  </p>
                  <p className="text-xs text-slate-400 mt-1">Định dạng Excel hỗ trợ: .xlsx, .xls</p>
                </div>
              </div>
            </div>

            {excelRows.length > 0 && (
              <div className="border border-emerald-100 bg-emerald-50/30 rounded-xl p-4 space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-xs font-bold text-emerald-800">BẢN XEM TRƯỚC HÀNG LOẠT ({excelRows.length} dòng)</span>
                  <button 
                    onClick={() => {
                      setExcelRows([]);
                      setExcelFileName("");
                      if (fileInputRef.current) fileInputRef.current.value = "";
                    }}
                    className="p-1 hover:bg-rose-100 rounded text-rose-600"
                    title="Xóa tệp chọn lại"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
                
                {/* Bảng xem trước dữ liệu thô */}
                <div className="max-h-40 overflow-y-auto border border-slate-200 rounded-lg bg-white text-[11px]">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-100 border-b border-slate-200 text-slate-600 font-bold sticky top-0">
                        <th className="p-2">STT</th>
                        {Object.keys(excelRows[0] || {}).slice(0, 3).map((colName) => (
                          <th key={colName} className="p-2 truncate max-w-[120px]">{colName}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="text-slate-600 divide-y divide-slate-100">
                      {excelRows.slice(0, 5).map((row, idx) => (
                        <tr key={idx}>
                          <td className="p-2 font-mono font-bold text-slate-400">{idx + 1}</td>
                          {Object.values(row).slice(0, 3).map((val: any, vIdx) => (
                            <td key={vIdx} className="p-2 truncate max-w-[120px]">{String(val)}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>

          <div className="mt-6">
            {uploadProgress ? (
              <div className="space-y-2">
                <div className="flex items-center gap-2 justify-center text-indigo-600 font-bold text-xs">
                  <Loader2 className="w-4 h-4 animate-spin" />
                  <span>{uploadProgress}</span>
                </div>
              </div>
            ) : (
              <button
                onClick={handleExcelUpload}
                disabled={loading || excelRows.length === 0}
                className="w-full py-2.5 bg-emerald-600 hover:bg-emerald-500 active:translate-y-0.5 transition font-bold text-white rounded-xl flex items-center justify-center gap-2 cursor-pointer text-sm shadow-md disabled:bg-slate-300"
              >
                <FileSpreadsheet className="w-4 h-4" />
                Tải lên & Khởi động lưu trữ Cloud
              </button>
            )}
          </div>
        </div>

      </div>

    </div>
  );
};
