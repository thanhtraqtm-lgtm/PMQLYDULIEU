import React, { useState, useEffect } from "react";
import { useAuth, db, isFirebaseInitialized } from "../context/AuthContext";
import { 
  collection, 
  query, 
  orderBy, 
  onSnapshot, 
  getDocs,
  limit
} from "firebase/firestore";
import { 
  Users, 
  Clock, 
  Activity, 
  ShieldCheck, 
  CheckCircle, 
  RefreshCw,
  Search,
  LayoutGrid
} from "lucide-react";

// Định nghĩa kiểu dữ liệu cho Nhật ký hành động
interface ActivityLog {
  id: string;
  unitID: string;
  userName: string;
  actionType: "manual_entry" | "excel_import";
  details: string;
  timestamp: string;
}

export const AdminDashboard: React.FC = () => {
  const [logs, setLogs] = useState<ActivityLog[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState("");

  // =======================================================
  // ĐỒNG BỘ THỜI GIAN THỰC (REAL-TIME FIRESTORE LISTENER)
  // =======================================================
  useEffect(() => {
    let unsubscribe: any = () => {};

    if (isFirebaseInitialized && db) {
      try {
        const logsRef = collection(db, "logs");
        const q = query(logsRef, orderBy("timestamp", "desc"), limit(200));

        // onSnapshot lắng nghe thay đổi dữ liệu từ Cloud Firestore và đồng bộ lập tức sang giao diện
        unsubscribe = onSnapshot(q, (snapshot) => {
          const fetchedLogs: ActivityLog[] = [];
          snapshot.forEach((doc) => {
            const data = doc.data();
            fetchedLogs.push({
              id: doc.id,
              unitID: data.unitID || "N/A",
              userName: data.userName || "Khách",
              actionType: data.actionType || "manual_entry",
              details: data.details || "",
              timestamp: data.timestamp || new Date().toISOString()
            });
          });
          setLogs(fetchedLogs);
          setLoading(false);
        }, (err) => {
          console.error("Lỗi lắng nghe Firestore logs:", err);
          setLoading(false);
        });

      } catch (err) {
        console.error("Lỗi thiết lập Firebase real-time listener:", err);
        setLoading(false);
      }
    } else {
      // Chế độ mô phỏng Offline: Đọc trực tiếp từ LocalStorage
      const fetchLocalLogs = () => {
        const raw = localStorage.getItem("system_mock_logs") || "[]";
        try {
          const parsed = JSON.parse(raw).map((l: any, idx: number) => ({
            id: `local_log_${idx}`,
            ...l
          }));
          setLogs(parsed);
        } catch {
          setLogs([]);
        }
        setLoading(false);
      };

      fetchLocalLogs();

      // Lắng nghe sự thay đổi của LocalStorage phát ra từ tab DataEntry để đồng bộ ngay tức khắc
      window.addEventListener("storage_logs_updated", fetchLocalLogs);
      return () => {
        window.removeEventListener("storage_logs_updated", fetchLocalLogs);
      };
    }

    return () => unsubscribe();
  }, []);

  // Định dạng hiển thị thời gian thân thiện với người Việt
  const formatTime = (isoString: string) => {
    try {
      const d = new Date(isoString);
      return d.toLocaleTimeString("vi-VN", {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit"
      }) + " " + d.toLocaleDateString("vi-VN", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric"
      });
    } catch {
      return isoString;
    }
  };

  // Tính toán số liệu thống kê nhanh từ logs thu thập được
  const uniqueUnits = Array.from(new Set(logs.map(l => l.unitID))).length;
  const manualCount = logs.filter(l => l.actionType === "manual_entry").length;
  const excelCount = logs.filter(l => l.actionType === "excel_import").length;

  // Lọc danh sách theo từ khóa tìm kiếm
  const filteredLogs = logs.filter(log => 
    log.unitID.toLowerCase().includes(searchTerm.toLowerCase()) ||
    log.details.toLowerCase().includes(searchTerm.toLowerCase()) ||
    log.userName.toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="space-y-6 font-sans">
      
      {/* Banner đầu trang Dashboard */}
      <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 p-5 bg-gradient-to-r from-slate-900 to-indigo-950 border border-slate-800 text-white rounded-2xl shadow-lg">
        <div>
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-6 h-6 text-indigo-400" />
            <h2 className="text-xl font-bold tracking-tight">Hệ Thống Kiểm Soát Hoạt Động Trung Tâm (Admin)</h2>
          </div>
          <p className="text-xs text-slate-300 mt-1">
            Theo dõi thao tác nạp và quản lý số liệu liên ngành từ tất cả các sở ban ngành trực thuộc thời gian thực.
          </p>
        </div>

        <div className="flex items-center gap-1.5 px-3.5 py-1.5 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded-full text-xs font-semibold">
          <Activity className="w-4 h-4 animate-pulse" />
          <span>Luồng Giám Sát LIVE</span>
        </div>
      </div>

      {/* Grid thẻ thống kê */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        
        <div className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm space-y-2">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-slate-400 uppercase">Số đơn vị gửi dữ liệu</span>
            <Users className="w-5 h-5 text-indigo-500" />
          </div>
          <div className="flex items-baseline gap-1">
            <span className="text-2xl font-bold text-slate-850">{uniqueUnits}</span>
            <span className="text-xs text-slate-500 font-medium">Cơ sở / Sở ban</span>
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm space-y-2">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-slate-400 uppercase">Khai báo thủ công</span>
            <span className="text-indigo-600 bg-indigo-50 p-1.5 rounded-lg text-xs font-bold font-mono">+{manualCount}</span>
          </div>
          <div className="flex items-baseline gap-1">
            <span className="text-2xl font-bold text-slate-850">{logs.filter(l => l.actionType === "manual_entry").length}</span>
            <span className="text-xs text-slate-500 font-medium">Bản ghi</span>
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm space-y-2">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-slate-400 uppercase">Tệp Excel đã nạp</span>
            <span className="text-emerald-600 bg-emerald-50 p-1.5 rounded-lg text-xs font-bold font-mono">+{excelCount}</span>
          </div>
          <div className="flex items-baseline gap-1">
            <span className="text-2xl font-bold text-slate-850">{logs.filter(l => l.actionType === "excel_import").length}</span>
            <span className="text-xs text-slate-500 font-medium">Phiên Import</span>
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm space-y-2">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-slate-400 uppercase">Trạng thái đồng bộ</span>
            <CheckCircle className="w-5 h-5 text-emerald-500" />
          </div>
          <div className="flex items-center gap-1.5">
            <span className="text-sm font-bold text-emerald-600">ỔN ĐỊNH 100%</span>
            <span className="text-xs text-slate-400">({isFirebaseInitialized ? "Firebase Active" : "Mock Active"})</span>
          </div>
        </div>

      </div>

      {/* Khối quản lý danh sách bảng log */}
      <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-4">
        
        {/* Thanh tìm kiếm và bộ lọc */}
        <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
          <h3 className="text-base font-bold text-slate-800 flex items-center gap-1.5">
            <Clock className="w-4 h-4 text-indigo-600" />
            Lịch sử thao tác liên ngành thời gian thực
          </h3>

          <div className="relative w-full sm:w-64">
            <span className="absolute inset-y-0 left-0 flex items-center pl-3">
              <Search className="w-4 h-4 text-slate-400" />
            </span>
            <input
              type="text"
              placeholder="Lọc theo đơn vị, nội dung..."
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
              className="w-full pl-9 pr-4 py-2 border border-slate-200 rounded-xl text-xs text-slate-700 bg-slate-50/50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500"
            />
          </div>
        </div>

        {/* Bảng dữ liệu 3 cột chính thức */}
        <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm">
          <table className="w-full text-left border-collapse text-xs text-slate-700">
            <thead>
              <tr className="bg-slate-100 border-b border-slate-200 text-slate-600 font-bold">
                {/* CỘT 1: Tên đơn vị */}
                <th className="p-3.5 w-1/4">TÊN ĐƠN VỊ (UNIT ID)</th>
                {/* CỘT 2: Mục thao tác */}
                <th className="p-3.5 w-2/4">MỤC THAO TÁC (ACTION DETAILS)</th>
                {/* CỘT 3: Thời gian thực */}
                <th className="p-3.5 w-1/4">THỜI GIAN THỰC (REAL-TIME TIMESTAMP)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-150">
              {loading ? (
                <tr>
                  <td colSpan={3} className="text-center p-8">
                    <div className="flex flex-col items-center justify-center space-y-2 text-slate-500">
                      <RefreshCw className="w-6 h-6 animate-spin text-indigo-500" />
                      <span className="font-semibold">Đang liên kết Cloud Real-time...</span>
                    </div>
                  </td>
                </tr>
              ) : filteredLogs.length === 0 ? (
                <tr>
                  <td colSpan={3} className="text-center p-12 text-slate-400">
                    Chưa có nhật ký hoạt động nào được ghi nhận. Hệ thống đang sẵn sàng!
                  </td>
                </tr>
              ) : (
                filteredLogs.map((log) => (
                  <tr key={log.id} className="hover:bg-slate-50/50 transition">
                    
                    {/* CỘT 1: Tên đơn vị */}
                    <td className="p-3.5">
                      <div className="flex items-center gap-2">
                        <span className="w-7 h-7 rounded-full bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-700 font-bold font-mono text-[10px]">
                          {log.unitID.slice(0, 2).toUpperCase()}
                        </span>
                        <div>
                          <span className="font-bold text-slate-800 uppercase font-mono block text-[11px]">{log.unitID}</span>
                          <span className="text-[10px] text-slate-400 font-medium">Nhân viên: {log.userName}</span>
                        </div>
                      </div>
                    </td>

                    {/* CỘT 2: Mục thao tác */}
                    <td className="p-3.5 leading-relaxed font-medium">
                      <div className="flex items-center gap-2">
                        <span className={`px-2 py-0.5 rounded-md text-[9px] font-bold ${
                          log.actionType === "excel_import" 
                            ? "bg-emerald-50 text-emerald-600 border border-emerald-100" 
                            : "bg-indigo-50 text-indigo-600 border border-indigo-100"
                        }`}>
                          {log.actionType === "excel_import" ? "Excel" : "Thủ công"}
                        </span>
                        <span className="text-slate-700 text-sm">{log.details}</span>
                      </div>
                    </td>

                    {/* CỘT 3: Thời gian thực */}
                    <td className="p-3.5 font-semibold text-slate-500 font-mono text-[11px]">
                      {formatTime(log.timestamp)}
                    </td>

                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

      </div>

    </div>
  );
};
