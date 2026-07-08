import React, { useState, useEffect } from "react";
import { 
  FileSpreadsheet, AlertCircle, RefreshCw, Layers, TrendingUp, Users, ArrowLeftRight, Sparkles, Building, Briefcase, ChevronRight, Minimize2, Maximize2 
} from "lucide-react";

interface SplitScreenViewProps {
  currentMode: "corp" | "individual";
  onSwitchMode: (mode: "corp" | "individual") => void;
  onRefreshAll?: () => void;
}

const DB_NAME = "VTongDatabase";
const STORE_NAME = "appState";
const CHUNK_SIZE = 50000;

export default function SplitScreenView({ currentMode, onSwitchMode, onRefreshAll }: SplitScreenViewProps) {
  const [corpMeta, setCorpMeta] = useState<any>(null);
  const [indMeta, setIndMeta] = useState<any>(null);
  
  const [corpRows, setCorpRows] = useState<any[]>([]);
  const [indRows, setIndRows] = useState<any[]>([]);
  
  const [loading, setLoading] = useState(false);
  const [statusMessage, setStatusMessage] = useState("");
  const [searchCorp, setSearchCorp] = useState("");
  const [searchInd, setSearchInd] = useState("");
  
  const [leftExpanded, setLeftExpanded] = useState(false);
  const [rightExpanded, setRightExpanded] = useState(false);

  // Helper to open DB
  const openDB = (): Promise<IDBDatabase> => {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, 1);
      request.onerror = () => reject(new Error("Không thể mở Database"));
      request.onsuccess = () => resolve(request.result);
    });
  };

  // Helper to load chunks using a single transaction for maximum speed
  const loadArrayInChunks = async (db: IDBDatabase, prefix: string, totalLength: number): Promise<any[]> => {
    const numChunks = Math.ceil(totalLength / CHUNK_SIZE);
    if (numChunks === 0) return [];
    
    const transaction = db.transaction(STORE_NAME, "readonly");
    const store = transaction.objectStore(STORE_NAME);
    
    const promises: Promise<any[]>[] = [];
    for (let i = 0; i < numChunks; i++) {
      promises.push(new Promise((resolve, reject) => {
        const request = store.get(`${prefix}_chunk_${i}`);
        request.onsuccess = () => resolve(request.result || []);
        request.onerror = () => reject(request.error || new Error(`Lỗi tải mảnh ${prefix} ${i}`));
      }));
    }
    
    const chunks = await Promise.all(promises);
    let result: any[] = [];
    for (let i = 0; i < chunks.length; i++) {
      result = result.concat(chunks[i]);
    }
    return result;
  };

  const loadData = async () => {
    setLoading(true);
    setStatusMessage("Đang quét và tải dữ liệu song song của cả hai hệ thống...");
    try {
      const db = await openDB();
      
      // Load Corp Metadata
      const cMeta: any = await new Promise((resolve) => {
        const transaction = db.transaction(STORE_NAME, "readonly");
        const store = transaction.objectStore(STORE_NAME);
        const request = store.get("sessionMeta_corp");
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => resolve(null);
      });
      setCorpMeta(cMeta);

      // Load Ind Metadata
      const iMeta: any = await new Promise((resolve) => {
        const transaction = db.transaction(STORE_NAME, "readonly");
        const store = transaction.objectStore(STORE_NAME);
        const request = store.get("sessionMeta_individual");
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => resolve(null);
      });
      setIndMeta(iMeta);

      // Load Corp Data
      if (cMeta && cMeta.mainDataLength) {
        setStatusMessage("Đang nạp dữ liệu Doanh nghiệp...");
        const cData = await loadArrayInChunks(db, "mainData_corp", cMeta.mainDataLength);
        setCorpRows(cData);
      } else {
        // Fallback to legacy single session if looking for corp on first load
        const legacyMeta: any = await new Promise((resolve) => {
          const transaction = db.transaction(STORE_NAME, "readonly");
          const store = transaction.objectStore(STORE_NAME);
          const request = store.get("sessionMeta");
          request.onsuccess = () => resolve(request.result || null);
          request.onerror = () => resolve(null);
        });
        if (legacyMeta && legacyMeta.mainDataLength) {
          const cData = await loadArrayInChunks(db, "mainData", legacyMeta.mainDataLength);
          setCorpRows(cData);
          setCorpMeta({
            fileName: legacyMeta.fileName,
            columns: legacyMeta.columns,
            mainDataLength: legacyMeta.mainDataLength
          });
        }
      }

      // Load Ind Data
      if (iMeta && iMeta.mainDataLength) {
        setStatusMessage("Đang nạp dữ liệu Hộ cá thể...");
        const iData = await loadArrayInChunks(db, "mainData_individual", iMeta.mainDataLength);
        setIndRows(iData);
      }
      
      setStatusMessage("");
    } catch (e: any) {
      console.error(e);
      setStatusMessage("Không phát hiện dữ liệu được lưu trữ riêng biệt.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, [currentMode]);

  // Statistics Calculation
  const calculateStats = (rows: any[], revenueCol: string, laborCol: string) => {
    if (!rows.length) return { total: 0, avgRev: 0, totalLabor: 0 };
    let totalRev = 0;
    let totalLabor = 0;
    let validRevCount = 0;
    let validLaborCount = 0;

    rows.forEach(r => {
      if (revenueCol && r[revenueCol] !== undefined) {
        const val = parseFloat(String(r[revenueCol]).replace(/,/g, ''));
        if (!isNaN(val)) {
          totalRev += val;
          validRevCount++;
        }
      }
      if (laborCol && r[laborCol] !== undefined) {
        const val = parseFloat(String(r[laborCol]).replace(/,/g, ''));
        if (!isNaN(val)) {
          totalLabor += val;
          validLaborCount++;
        }
      }
    });

    return {
      total: rows.length,
      avgRev: validRevCount > 0 ? Math.round(totalRev / validRevCount) : 0,
      totalLabor: totalLabor
    };
  };

  const corpStats = calculateStats(
    corpRows, 
    corpMeta?.mapping?.doanhthu || "", 
    corpMeta?.mapping?.laodong || ""
  );

  const indStats = calculateStats(
    indRows, 
    indMeta?.mapping?.doanhthu || "", 
    indMeta?.mapping?.laodong || ""
  );

  // Filtering
  const filteredCorpRows = corpRows.filter(r => {
    if (!searchCorp) return true;
    return Object.values(r).some(val => 
      String(val).toLowerCase().includes(searchCorp.toLowerCase())
    );
  }).slice(0, 50);

  const filteredIndRows = indRows.filter(r => {
    if (!searchInd) return true;
    return Object.values(r).some(val => 
      String(val).toLowerCase().includes(searchInd.toLowerCase())
    );
  }).slice(0, 50);

  return (
    <div className="space-y-6 animate-fade-in font-sans">
      
      {/* HEADER COGNITIVE PANEL */}
      <div className="bg-white border border-slate-200 rounded-3xl p-5 shadow-sm space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div>
            <span className="bg-indigo-50 text-indigo-700 text-[10px] font-extrabold uppercase tracking-widest px-3 py-1 rounded-full border border-indigo-100">
              Màn Hình Song Song
            </span>
            <h3 className="text-lg font-black text-slate-800 tracking-tight mt-1.5 flex items-center gap-2">
              <ArrowLeftRight className="w-5 h-5 text-indigo-500 shrink-0" />
              Rà Soát Song Song 50/50: Doanh Nghiệp vs Cá Thể
            </h3>
            <p className="text-xs text-slate-500">
              Xem trực quan, so sánh quy mô và tương tác trực tiếp trên cả hai luồng dữ liệu độc lập không liên quan đến nhau.
            </p>
          </div>
          
          <button
            onClick={loadData}
            className="bg-slate-50 border border-slate-200 text-slate-600 hover:bg-indigo-50 hover:text-indigo-600 hover:border-indigo-200 text-xs font-bold px-4 py-2 rounded-xl transition-all flex items-center gap-1.5 cursor-pointer shadow-sm"
          >
            <RefreshCw className={`w-4 h-4 ${loading ? "animate-spin" : ""}`} /> Làm mới dữ liệu
          </button>
        </div>

        {statusMessage && (
          <div className="bg-indigo-50/50 text-indigo-700 border border-indigo-100 p-3 rounded-xl text-xs font-medium animate-pulse">
            ⚡ {statusMessage}
          </div>
        )}
      </div>

      {/* 50/50 SPLIT SCREEN LAYOUT */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 items-stretch">
        
        {/* DOANH NGHIỆP (LEFT PANE) */}
        {(!rightExpanded) && (
          <div className={`bg-white border-2 rounded-3xl p-5 shadow-sm flex flex-col justify-between transition-all duration-300 ${
            currentMode === "corp" ? "border-indigo-500 shadow-indigo-100 shadow-md" : "border-slate-200"
          } ${leftExpanded ? "col-span-2" : ""}`}>
            
            <div className="space-y-4">
              {/* Header */}
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2.5">
                  <div className="p-2.5 bg-indigo-50 text-indigo-600 rounded-xl border border-indigo-100">
                    <Building className="w-5 h-5" />
                  </div>
                  <div>
                    <h4 className="text-sm font-black text-slate-800 uppercase tracking-wide">Màn 1: Doanh Nghiệp (Corp)</h4>
                    <p className="text-[11px] text-slate-500 font-mono">
                      {corpMeta?.fileName ? `📁 ${corpMeta.fileName}` : "❌ Chưa nạp file dữ liệu doanh nghiệp"}
                    </p>
                  </div>
                </div>
                
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => {
                      setLeftExpanded(!leftExpanded);
                      setRightExpanded(false);
                    }}
                    className="p-1.5 hover:bg-slate-100 rounded text-slate-500 hover:text-slate-700 cursor-pointer"
                    title={leftExpanded ? "Thu nhỏ" : "Phóng to toàn màn hình"}
                  >
                    {leftExpanded ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
                  </button>
                  <button
                    onClick={() => onSwitchMode("corp")}
                    disabled={currentMode === "corp"}
                    className={`px-3 py-1.5 rounded-lg text-xs font-extrabold cursor-pointer transition-all ${
                      currentMode === "corp" 
                        ? "bg-indigo-600 text-white shadow-sm cursor-default" 
                        : "bg-slate-100 hover:bg-indigo-50 text-slate-600 hover:text-indigo-600"
                    }`}
                  >
                    {currentMode === "corp" ? "Đang chọn" : "Kích hoạt thao tác ➔"}
                  </button>
                </div>
              </div>

              {/* Stats Bar */}
              <div className="grid grid-cols-3 gap-3 bg-slate-50 p-3 rounded-2xl border border-slate-150">
                <div className="text-center">
                  <span className="text-[10px] text-slate-400 font-bold font-mono block">QUY MÔ</span>
                  <strong className="text-base text-slate-800 font-mono font-black">
                    {corpStats.total.toLocaleString()}
                  </strong>
                  <span className="text-[9px] text-slate-400 block">Doanh nghiệp</span>
                </div>
                <div className="text-center border-x border-slate-200">
                  <span className="text-[10px] text-slate-400 font-bold font-mono block">DTHU BÌNH QUÂN</span>
                  <strong className="text-base text-emerald-600 font-mono font-black">
                    {corpStats.avgRev.toLocaleString()}
                  </strong>
                  <span className="text-[9px] text-slate-400 block">Tr.đồng / Đơn vị</span>
                </div>
                <div className="text-center">
                  <span className="text-[10px] text-slate-400 font-bold font-mono block">TỔNG LAO ĐỘNG</span>
                  <strong className="text-base text-indigo-600 font-mono font-black">
                    {corpStats.totalLabor.toLocaleString()}
                  </strong>
                  <span className="text-[9px] text-slate-400 block">Nhân sự thực tế</span>
                </div>
              </div>

              {/* Table Data */}
              {corpRows.length > 0 ? (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <input 
                      type="text"
                      value={searchCorp}
                      onChange={(e) => setSearchCorp(e.target.value)}
                      placeholder="Tìm kiếm nhanh dòng doanh nghiệp..."
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 text-xs focus:outline-none focus:ring-1 focus:ring-indigo-500 font-sans shadow-inner"
                    />
                  </div>

                  <div className="overflow-x-auto max-h-[300px] border border-slate-200 rounded-2xl bg-white shadow-inner">
                    <table className="w-full text-left text-xs min-w-[500px]">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-mono text-[11px] sticky top-0 z-10">
                          {corpMeta?.columns?.slice(0, 5).map((col: string) => (
                            <th key={col} className="p-2.5 font-bold border-b border-slate-200 bg-slate-50">{col}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-sans">
                        {filteredCorpRows.map((row, idx) => (
                          <tr key={idx} className="hover:bg-indigo-50/20 transition-colors">
                            {corpMeta?.columns?.slice(0, 5).map((col: string) => (
                              <td key={col} className="p-2.5 text-slate-700 max-w-[150px] truncate" title={row[col]}>
                                {row[col] !== undefined ? String(row[col]) : "-"}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <p className="text-[10px] text-slate-400 italic text-right">
                    * Đang xem nhanh 50 dòng đầu tiên của bảng Doanh Nghiệp.
                  </p>
                </div>
              ) : (
                <div className="border-2 border-dashed border-slate-200 rounded-2xl p-8 text-center text-slate-400 text-xs flex flex-col items-center justify-center space-y-2">
                  <FileSpreadsheet className="w-8 h-8 text-slate-300" />
                  <span>Chưa nạp dữ liệu Doanh nghiệp.</span>
                  <button 
                    onClick={() => {
                      onSwitchMode("corp");
                      onRefreshAll && onRefreshAll();
                    }}
                    className="mt-2 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-700 font-bold text-[11px] px-3 py-1.5 rounded-lg transition-all cursor-pointer"
                  >
                    📂 Chuyển Sang Tab Nạp Doanh Nghiệp
                  </button>
                </div>
              )}
            </div>
            
          </div>
        )}

        {/* HỘ CÁ THỂ (RIGHT PANE) */}
        {(!leftExpanded) && (
          <div className={`bg-white border-2 rounded-3xl p-5 shadow-sm flex flex-col justify-between transition-all duration-300 ${
            currentMode === "individual" ? "border-indigo-500 shadow-indigo-100 shadow-md" : "border-slate-200"
          } ${rightExpanded ? "col-span-2" : ""}`}>
            
            <div className="space-y-4">
              {/* Header */}
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2.5">
                  <div className="p-2.5 bg-orange-50 text-orange-600 rounded-xl border border-orange-100">
                    <Briefcase className="w-5 h-5" />
                  </div>
                  <div>
                    <h4 className="text-sm font-black text-slate-800 uppercase tracking-wide">Màn 2: Hộ Cá Thể (Individual)</h4>
                    <p className="text-[11px] text-slate-500 font-mono">
                      {indMeta?.fileName ? `📁 ${indMeta.fileName}` : "❌ Chưa nạp file dữ liệu hộ cá thể"}
                    </p>
                  </div>
                </div>
                
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => {
                      setRightExpanded(!rightExpanded);
                      setLeftExpanded(false);
                    }}
                    className="p-1.5 hover:bg-slate-100 rounded text-slate-500 hover:text-slate-700 cursor-pointer"
                    title={rightExpanded ? "Thu nhỏ" : "Phóng to toàn màn hình"}
                  >
                    {rightExpanded ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
                  </button>
                  <button
                    onClick={() => onSwitchMode("individual")}
                    disabled={currentMode === "individual"}
                    className={`px-3 py-1.5 rounded-lg text-xs font-extrabold cursor-pointer transition-all ${
                      currentMode === "individual" 
                        ? "bg-indigo-600 text-white shadow-sm cursor-default" 
                        : "bg-slate-100 hover:bg-indigo-50 text-slate-600 hover:text-indigo-600"
                    }`}
                  >
                    {currentMode === "individual" ? "Đang chọn" : "Kích hoạt thao tác ➔"}
                  </button>
                </div>
              </div>

              {/* Stats Bar */}
              <div className="grid grid-cols-3 gap-3 bg-slate-50 p-3 rounded-2xl border border-slate-150">
                <div className="text-center">
                  <span className="text-[10px] text-slate-400 font-bold font-mono block">QUY MÔ</span>
                  <strong className="text-base text-slate-800 font-mono font-black">
                    {indStats.total.toLocaleString()}
                  </strong>
                  <span className="text-[9px] text-slate-400 block">Hộ cá thể</span>
                </div>
                <div className="text-center border-x border-slate-200">
                  <span className="text-[10px] text-slate-400 font-bold font-mono block">DTHU BÌNH QUÂN</span>
                  <strong className="text-base text-emerald-600 font-mono font-black">
                    {indStats.avgRev.toLocaleString()}
                  </strong>
                  <span className="text-[9px] text-slate-400 block">Tr.đồng / Hộ</span>
                </div>
                <div className="text-center">
                  <span className="text-[10px] text-slate-400 font-bold font-mono block">TỔNG LAO ĐỘNG</span>
                  <strong className="text-base text-indigo-600 font-mono font-black">
                    {indStats.totalLabor.toLocaleString()}
                  </strong>
                  <span className="text-[9px] text-slate-400 block">Nhân sự thực tế</span>
                </div>
              </div>

              {/* Table Data */}
              {indRows.length > 0 ? (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <input 
                      type="text"
                      value={searchInd}
                      onChange={(e) => setSearchInd(e.target.value)}
                      placeholder="Tìm kiếm nhanh dòng hộ cá thể..."
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 text-xs focus:outline-none focus:ring-1 focus:ring-indigo-500 font-sans shadow-inner"
                    />
                  </div>

                  <div className="overflow-x-auto max-h-[300px] border border-slate-200 rounded-2xl bg-white shadow-inner">
                    <table className="w-full text-left text-xs min-w-[500px]">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-mono text-[11px] sticky top-0 z-10">
                          {indMeta?.columns?.slice(0, 5).map((col: string) => (
                            <th key={col} className="p-2.5 font-bold border-b border-slate-200 bg-slate-50">{col}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-sans">
                        {filteredIndRows.map((row, idx) => (
                          <tr key={idx} className="hover:bg-indigo-50/20 transition-colors">
                            {indMeta?.columns?.slice(0, 5).map((col: string) => (
                              <td key={col} className="p-2.5 text-slate-700 max-w-[150px] truncate" title={row[col]}>
                                {row[col] !== undefined ? String(row[col]) : "-"}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <p className="text-[10px] text-slate-400 italic text-right">
                    * Đang xem nhanh 50 dòng đầu tiên của bảng Hộ Cá Thể.
                  </p>
                </div>
              ) : (
                <div className="border-2 border-dashed border-slate-200 rounded-2xl p-8 text-center text-slate-400 text-xs flex flex-col items-center justify-center space-y-2">
                  <Briefcase className="w-8 h-8 text-slate-300" />
                  <span>Chưa nạp dữ liệu Hộ cá thể.</span>
                  <button 
                    onClick={() => {
                      onSwitchMode("individual");
                      onRefreshAll && onRefreshAll();
                    }}
                    className="mt-2 bg-orange-50 hover:bg-orange-100 border border-orange-200 text-orange-700 font-bold text-[11px] px-3 py-1.5 rounded-lg transition-all cursor-pointer"
                  >
                    📂 Chuyển Sang Tab Nạp Cá Thể
                  </button>
                </div>
              )}
            </div>

          </div>
        )}

      </div>

      {/* QUICK COMPARISON BOX */}
      {corpRows.length > 0 && indRows.length > 0 && (
        <div className="bg-indigo-900 text-white rounded-3xl p-6 shadow-xl relative overflow-hidden">
          <div className="absolute top-0 right-0 w-32 h-32 bg-indigo-500/10 rounded-full blur-3xl opacity-60"></div>
          
          <div className="flex flex-col md:flex-row items-center justify-between gap-4">
            <div>
              <span className="bg-indigo-800 text-indigo-200 text-[9px] font-black tracking-widest uppercase px-3 py-1 rounded-full">
                Phân tích đối sánh
              </span>
              <h4 className="text-sm font-black mt-1.5 flex items-center gap-1.5 uppercase">
                <Sparkles className="w-4 h-4 text-amber-400" /> So sánh quy mô Kinh tế đa tầng
              </h4>
              <p className="text-xs text-indigo-200 mt-1">
                Dữ liệu Doanh Nghiệp gấp <b>{(corpStats.total / (indStats.total || 1)).toFixed(1)} lần</b> Hộ Cá Thể về số lượng đơn vị.
              </p>
            </div>
            
            <div className="flex gap-4">
              <div className="bg-indigo-950/40 border border-indigo-800 p-3 rounded-xl text-center">
                <span className="text-[10px] text-indigo-300 block font-mono">TỔNG THU NHẬP DN</span>
                <span className="text-sm font-extrabold text-white font-mono">
                  {(corpStats.total * corpStats.avgRev).toLocaleString()} tr.đ
                </span>
              </div>
              <div className="bg-indigo-950/40 border border-indigo-800 p-3 rounded-xl text-center">
                <span className="text-[10px] text-indigo-300 block font-mono">TỔNG THU NHẬP HỘ</span>
                <span className="text-sm font-extrabold text-white font-mono">
                  {(indStats.total * indStats.avgRev).toLocaleString()} tr.đ
                </span>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
