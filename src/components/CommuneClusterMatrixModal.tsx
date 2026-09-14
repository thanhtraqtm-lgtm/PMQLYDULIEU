import React, { useState, useMemo } from "react";
import {
  Building2,
  CheckCircle2,
  AlertTriangle,
  ArrowRight,
  Plus,
  Trash2,
  Download,
  Search,
  RefreshCw,
  X,
  Layers,
  Sparkles,
  Play
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  OFFICIAL_COMMUNE_MERGE_DATA,
  cleanCommuneStr,
  findOfficialNewCommune,
  findOfficialByCommuneCode,
  findOfficialNewCommuneCode,
} from "../utils/villageAreaCrosswalk";

export interface CommuneClusterMatrixModalProps {
  dataA: any[];
  colCommuneNameA: string;
  colCommuneCodeA?: string;
  colVillageNameA: string;
  dataB: any[];
  colCommuneNameB: string;
  colCommuneCodeB?: string;
  colVillageNameB: string;
  customCommuneCrosswalk: Array<{ communeA: string; communeB: string }>;
  onUpdateCustomCrosswalk: (newList: Array<{ communeA: string; communeB: string }>) => void;
  onSelectCommuneToFilter: (communeAName: string) => void;
  onClose: () => void;
}

export const CommuneClusterMatrixModal: React.FC<CommuneClusterMatrixModalProps> = ({
  dataA,
  colCommuneNameA,
  colCommuneCodeA,
  colVillageNameA,
  dataB,
  colCommuneNameB,
  colCommuneCodeB,
  colVillageNameB,
  customCommuneCrosswalk,
  onUpdateCustomCrosswalk,
  onSelectCommuneToFilter,
  onClose,
}) => {
  const [searchTerm, setSearchTerm] = useState("");
  const [filterStatus, setFilterStatus] = useState<"ALL" | "HAS_OLD" | "EMPTY" | "UNBALANCED">("ALL");
  const [addingToNewCommune, setAddingToNewCommune] = useState<string | null>(null);
  const [selectedOldCommuneToAdd, setSelectedOldCommuneToAdd] = useState<string>("");

  // 1. Trích xuất danh sách tất cả các XÃ CŨ duy nhất trong Tệp A
  const oldCommunesList = useMemo(() => {
    if (!colCommuneNameA || dataA.length === 0) return [];
    const map = new Map<string, { name: string; code: string; villageCount: number }>();
    
    dataA.forEach(r => {
      const name = String(r[colCommuneNameA] || "").trim();
      const code = colCommuneCodeA ? String(r[colCommuneCodeA] || "").trim() : "";
      if (name) {
        const existing = map.get(name) || { name, code, villageCount: 0 };
        existing.villageCount += 1;
        if (!existing.code && code) existing.code = code;
        map.set(name, existing);
      }
    });

    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name, "vi"));
  }, [dataA, colCommuneNameA, colCommuneCodeA]);

  // 2. Trích xuất danh sách tất cả các XÃ MỚI duy nhất trong Tệp B (Ví dụ đúng 65 xã mới)
  const newCommunesList = useMemo(() => {
    if (!colCommuneNameB || dataB.length === 0) return [];
    const map = new Map<string, { name: string; code: string; villageCount: number }>();

    dataB.forEach(r => {
      const name = String(r[colCommuneNameB] || "").trim();
      const code = colCommuneCodeB ? String(r[colCommuneCodeB] || "").trim() : "";
      if (name) {
        const existing = map.get(name) || { name, code, villageCount: 0 };
        existing.villageCount += 1;
        if (!existing.code && code) existing.code = code;
        map.set(name, existing);
      }
    });

    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name, "vi"));
  }, [dataB, colCommuneNameB, colCommuneCodeB]);

  // Map ánh xạ Xã cũ ➔ Xã mới
  // Ưu tiên: 1. Cấu hình tay -> 2. Khớp Mã Xã Cũ (codeA ➔ codeB/communeB) -> 3. Từ điển chính thức -> 4. Trùng tên trực tiếp
  const oldToNewMap = useMemo(() => {
    const map = new Map<string, string>(); // oldCommuneName -> newCommuneName

    oldCommunesList.forEach(oldC => {
      // 1. Kiểm tra bằng Mã Xã Cũ (codeA) trong OFFICIAL_COMMUNE_MERGE_DATA
      if (oldC.code) {
        const byCode = findOfficialByCommuneCode(oldC.code);
        if (byCode) {
          const matchInNew = newCommunesList.find(
            newC =>
              (newC.code && String(newC.code).trim() === String(byCode.codeB).trim()) ||
              cleanCommuneStr(newC.name) === cleanCommuneStr(byCode.communeB)
          );
          if (matchInNew) {
            map.set(oldC.name, matchInNew.name);
            return;
          }
        }
      }

      // 2. Tra cứu trong từ điển chính thức bằng Tên xã cũ
      const officialNew = findOfficialNewCommune(oldC.name);
      if (officialNew) {
        const matchInNew = newCommunesList.find(
          newC => cleanCommuneStr(newC.name) === cleanCommuneStr(officialNew)
        );
        if (matchInNew) {
          map.set(oldC.name, matchInNew.name);
          return;
        }
      }

      // 3. Tự động map trùng tên trực tiếp
      const cleanOld = cleanCommuneStr(oldC.name);
      const exactMatch = newCommunesList.find(newC => cleanCommuneStr(newC.name) === cleanOld);
      if (exactMatch) {
        map.set(oldC.name, exactMatch.name);
      }
    });

    // Sau cùng đè cấu hình tay của người dùng lên
    customCommuneCrosswalk.forEach(c => {
      if (c.communeA && c.communeB) {
        map.set(c.communeA, c.communeB);
      }
    });

    return map;
  }, [oldCommunesList, newCommunesList, customCommuneCrosswalk]);

  // 3. Gom cụm cho các Xã Mới
  const clusterMatrix = useMemo(() => {
    return newCommunesList.map(newC => {
      // Tìm tất cả các xã cũ đã trỏ về xã mới này
      const mappedOlds: Array<{ name: string; code: string; villageCount: number }> = [];

      oldCommunesList.forEach(oldC => {
        const targetNew = oldToNewMap.get(oldC.name);
        if (targetNew && cleanCommuneStr(targetNew) === cleanCommuneStr(newC.name)) {
          mappedOlds.push(oldC);
        }
      });

      const totalOldVillages = mappedOlds.reduce((sum, item) => sum + item.villageCount, 0);
      const diffVillages = newC.villageCount - totalOldVillages;

      return {
        newCommune: newC,
        mappedOlds,
        totalOldVillages,
        diffVillages,
      };
    });
  }, [newCommunesList, oldCommunesList, oldToNewMap]);

  // 4. Danh sách các XÃ CŨ CHƯA ĐƯỢC GOM VÀO XÃ MỚI NÀO (Mồ côi)
  const unmappedOldCommunes = useMemo(() => {
    return oldCommunesList.filter(oldC => !oldToNewMap.has(oldC.name));
  }, [oldCommunesList, oldToNewMap]);

  // Tự động gom nhanh tất cả xã cũ vào xã mới dựa trên từ điển sáp nhập
  const handleAutoGroupAll = () => {
    const newCustoms = [...customCommuneCrosswalk];
    let addedCount = 0;

    unmappedOldCommunes.forEach(oldC => {
      // 1. Kiểm tra bằng Mã Xã Cũ
      if (oldC.code) {
        const byCode = findOfficialByCommuneCode(oldC.code);
        if (byCode) {
          const match = newCommunesList.find(
            n =>
              (n.code && String(n.code).trim() === String(byCode.codeB).trim()) ||
              cleanCommuneStr(n.name) === cleanCommuneStr(byCode.communeB)
          );
          if (match) {
            newCustoms.push({ communeA: oldC.name, communeB: match.name });
            addedCount++;
            return;
          }
        }
      }

      // 2. Tra cứu bằng Tên xã cũ
      const officialNew = findOfficialNewCommune(oldC.name);
      if (officialNew) {
        const match = newCommunesList.find(n => cleanCommuneStr(n.name) === cleanCommuneStr(officialNew));
        if (match) {
          newCustoms.push({ communeA: oldC.name, communeB: match.name });
          addedCount++;
          return;
        }
      }

      // 3. Thử tìm theo tên tương đồng
      const cleanOld = cleanCommuneStr(oldC.name);
      const match = newCommunesList.find(n => cleanCommuneStr(n.name) === cleanOld);
      if (match) {
        newCustoms.push({ communeA: oldC.name, communeB: match.name });
        addedCount++;
      }
    });

    if (addedCount > 0) {
      onUpdateCustomCrosswalk(newCustoms);
      alert(`Đã tự động nhận diện và ghép thêm ${addedCount} xã cũ vào đúng các xã mới tương ứng!`);
    } else {
      alert("Không tìm thêm được cặp xã tự động nào. Bạn có thể chọn ghép thủ công bằng nút [+ Thêm xã cũ].");
    }
  };

  // Xuất file Excel Từ Điển Chuẩn Toàn Diện (Mã xã cũ ➔ Mã xã mới)
  const handleExportOfficialCrosswalk = () => {
    const exportRows = OFFICIAL_COMMUNE_MERGE_DATA.map((r, idx) => ({
      STT: idx + 1,
      "Mã xã cũ": r.codeA || "",
      "Tên Xã/Phường Cũ": r.communeA,
      "Tên Xã/Phường Mới": r.communeB,
      "Mã xã mới": r.codeB || "",
      "Ghi chú / Huyện": r.districtNote || "",
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tu_Dien_Xa_Cu_Moi_Chuan");
    XLSX.writeFile(wb, "Bang_Anh_Xa_Chuan_Ma_Xa_Cu_Sang_Xa_Moi_104_Xa.xlsx");
  };

  // Thêm thủ công 1 xã cũ vào 1 xã mới
  const handleAddOldToNew = (newCommuneName: string) => {
    if (!selectedOldCommuneToAdd) return;
    const filtered = customCommuneCrosswalk.filter(c => c.communeA !== selectedOldCommuneToAdd);
    filtered.push({ communeA: selectedOldCommuneToAdd, communeB: newCommuneName });
    onUpdateCustomCrosswalk(filtered);
    setAddingToNewCommune(null);
    setSelectedOldCommuneToAdd("");
  };

  // Gỡ 1 xã cũ ra khỏi xã mới
  const handleRemoveOldFromNew = (oldCommuneName: string) => {
    const filtered = customCommuneCrosswalk.filter(c => c.communeA !== oldCommuneName);
    onUpdateCustomCrosswalk(filtered);
  };

  // Xuất file Excel Ma Trận Gom Cụm Xã
  const handleExportClusterExcel = () => {
    const exportRows = clusterMatrix.map((item, idx) => ({
      STT: idx + 1,
      "Tên Xã/Phường Mới": item.newCommune.name,
      "Mã Xã Mới": item.newCommune.code || "",
      "Số ĐBĐT Mới (Kỳ này)": item.newCommune.villageCount,
      "Số Xã Cũ Gộp Về": item.mappedOlds.length,
      "Danh Sách Xã Cũ Sáp Nhập": item.mappedOlds.map(o => `${o.name} (${o.villageCount} địa bàn)`).join("; ") || "Chưa có xã cũ nào",
      "Tổng Địa Bàn Cũ Gom Về": item.totalOldVillages,
      "Chênh Lệch Địa Bàn (Mới - Cũ)": item.diffVillages,
      "Đánh Giá":
        item.mappedOlds.length === 0
          ? "Chưa ghép xã cũ"
          : item.totalOldVillages === item.newCommune.villageCount
          ? "Đủ và khớp bằng số địa bàn"
          : item.diffVillages > 0
          ? `Tách thêm ${item.diffVillages} ĐBĐT mới`
          : `Gộp bớt ${Math.abs(item.diffVillages)} địa bàn`,
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Ma_Tran_65_Xa_Moi");
    XLSX.writeFile(wb, `Ma_Tran_Gom_${newCommunesList.length}_Xa_Moi_Du_Dia_Ban.xlsx`);
  };

  // Lọc hiển thị
  const filteredClusters = useMemo(() => {
    return clusterMatrix.filter(c => {
      if (filterStatus === "HAS_OLD" && c.mappedOlds.length === 0) return false;
      if (filterStatus === "EMPTY" && c.mappedOlds.length > 0) return false;
      if (filterStatus === "UNBALANCED" && c.mappedOlds.length > 0 && c.totalOldVillages === c.newCommune.villageCount) return false;

      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase().trim();
        const matchNew = c.newCommune.name.toLowerCase().includes(q) || c.newCommune.code.includes(q);
        const matchOld = c.mappedOlds.some(o => o.name.toLowerCase().includes(q) || o.code.includes(q));
        return matchNew || matchOld;
      }

      return true;
    });
  }, [clusterMatrix, filterStatus, searchTerm]);

  return (
    <div className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-3 sm:p-6 backdrop-blur-xs">
      <div className="bg-white w-full max-w-6xl max-h-[92vh] flex flex-col border-2 border-emerald-700 shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-150">
        {/* HEADER MODAL */}
        <div className="bg-emerald-800 text-white p-3.5 sm:p-4 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2.5">
            <div className="w-10 h-10 bg-white/10 flex items-center justify-center border border-white/20">
              <Layers className="w-6 h-6 text-amber-300" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-base font-bold uppercase tracking-wider text-white">
                  MA TRẬN GOM CỤM XÃ: {newCommunesList.length} XÃ MỚI ➔ TÌM ĐỦ TOÀN BỘ XÃ CŨ &amp; ĐỊA BÀN
                </h3>
                <span className="bg-amber-400 text-slate-900 text-[10px] font-black px-2 py-0.5 uppercase">
                  Bắt Từ Cấp Xã
                </span>
              </div>
              <p className="text-xs text-emerald-100 m-0">
                Gom đúng và đủ các xã cũ vào từng xã mới. Khi đã đủ cụm xã thì toàn bộ các thôn/địa bàn bên trong sẽ đủ 100%, không bị rơi rụng.
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleExportOfficialCrosswalk}
              className="bg-amber-400 hover:bg-amber-500 text-slate-900 text-xs font-black px-2.5 py-1.5 flex items-center gap-1 cursor-pointer shadow-xs"
              title="Xuất bảng đối soát chuẩn Mã xã cũ ➔ Mã xã mới (Đủ 104 xã sáp nhập)"
            >
              <Download className="w-4 h-4 text-slate-900" />
              <span>Xuất mã xã cũ - mới</span>
            </button>

            <button
              type="button"
              onClick={handleExportClusterExcel}
              className="bg-emerald-950 hover:bg-emerald-900 text-white text-xs font-bold px-2.5 py-1.5 flex items-center gap-1 border border-emerald-500/50 cursor-pointer shadow-xs"
              title="Xuất bảng gom xã mới kèm danh sách xã cũ và số lượng địa bàn ra Excel"
            >
              <Download className="w-4 h-4 text-emerald-300" />
              <span>Xuất gom xã</span>
            </button>

            <button
              type="button"
              onClick={onClose}
              className="text-white hover:bg-white/20 p-1.5 cursor-pointer border-0 bg-transparent"
              title="Đóng cửa sổ"
            >
              <X className="w-6 h-6" />
            </button>
          </div>
        </div>

        {/* THÔNG KÊ NHANH & CẢNH BÁO XÃ CŨ MỒ CÔI */}
        <div className="p-3 sm:p-4 bg-slate-50 border-b border-slate-300 shrink-0 space-y-3">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
            <div className="bg-white p-2.5 border border-slate-300 shadow-2xs">
              <span className="text-slate-500 font-bold block text-[11px]">TỔNG XÃ MỚI (TỆP B):</span>
              <span className="text-xl font-black text-emerald-800">{newCommunesList.length} Xã/Phường</span>
              <span className="text-[11px] text-slate-500 block">({dataB.length.toLocaleString("vi-VN")} địa bàn mới)</span>
            </div>

            <div className="bg-white p-2.5 border border-slate-300 shadow-2xs">
              <span className="text-slate-500 font-bold block text-[11px]">TỔNG XÃ CŨ (TỆP A):</span>
              <span className="text-xl font-black text-sky-800">{oldCommunesList.length} Xã/Thị trấn</span>
              <span className="text-[11px] text-slate-500 block">({dataA.length.toLocaleString("vi-VN")} địa bàn cũ)</span>
            </div>

            <div className="bg-white p-2.5 border border-slate-300 shadow-2xs">
              <span className="text-slate-500 font-bold block text-[11px]">XÃ CŨ ĐÃ ĐƯỢC GOM:</span>
              <span className="text-xl font-black text-emerald-700">
                {oldCommunesList.length - unmappedOldCommunes.length} / {oldCommunesList.length}
              </span>
              <span className="text-[11px] text-emerald-700 font-medium block">
                (Đạt {oldCommunesList.length > 0 ? (((oldCommunesList.length - unmappedOldCommunes.length) / oldCommunesList.length) * 100).toFixed(0) : 0}%)
              </span>
            </div>

            <div className={`p-2.5 border shadow-2xs ${unmappedOldCommunes.length > 0 ? "bg-rose-50 border-rose-300 text-rose-900" : "bg-emerald-50 border-emerald-300 text-emerald-900"}`}>
              <span className="font-bold block text-[11px]">XÃ CŨ CHƯA ĐƯỢC GOM:</span>
              <span className="text-xl font-black">
                {unmappedOldCommunes.length} Xã
              </span>
              <span className="text-[11px] font-medium block">
                {unmappedOldCommunes.length === 0 ? "✅ Đã gom đủ 100% xã cũ" : "⚠️ Cần gom nốt vào các xã mới"}
              </span>
            </div>
          </div>

          {/* CẢNH BÁO VÀ HỘP DANH SÁCH XÃ CŨ CHƯA GOM */}
          {unmappedOldCommunes.length > 0 && (
            <div className="bg-amber-50 border border-amber-300 p-2.5 space-y-2">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-1.5 text-xs text-amber-950 font-bold">
                  <AlertTriangle className="w-4 h-4 text-amber-700 shrink-0" />
                  <span>
                    Còn {unmappedOldCommunes.length} xã cũ chưa được xếp vào xã mới nào (gồm tổng {unmappedOldCommunes.reduce((s, c) => s + c.villageCount, 0)} địa bàn cũ):
                  </span>
                </div>

                <button
                  type="button"
                  onClick={handleAutoGroupAll}
                  className="bg-amber-700 hover:bg-amber-800 text-white text-xs font-bold px-3 py-1 flex items-center gap-1.5 cursor-pointer border-0 shadow-2xs"
                  title="Tự động tra cứu trong Từ điển 310 xã sáp nhập chính thức để ghép nhanh"
                >
                  <Sparkles className="w-3.5 h-3.5 text-amber-200" />
                  <span>⚡ Tự Động Gom Nhanh Theo Từ Điển Sáp Nhập</span>
                </button>
              </div>

              <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto p-1 bg-white border border-amber-200">
                {unmappedOldCommunes.map(c => (
                  <span
                    key={c.name}
                    className="inline-flex items-center gap-1 text-[11px] bg-rose-100 text-rose-900 px-2 py-0.5 border border-rose-300 font-bold"
                  >
                    <span>{c.name}</span>
                    <span className="text-[10px] text-rose-700 font-mono">({c.villageCount} đb)</span>
                  </span>
                ))}
              </div>
            </div>
          )}

          {/* THANH TÌM KIẾM & BỘ LỌC CỤM XÃ */}
          <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
            <div className="flex items-center gap-2 flex-1 min-w-[280px]">
              <div className="relative flex-1">
                <Search className="w-4 h-4 absolute left-2.5 top-2 text-slate-400" />
                <input
                  type="text"
                  value={searchTerm}
                  onChange={e => setSearchTerm(e.target.value)}
                  placeholder="Tìm theo tên xã mới hoặc xã cũ sáp nhập..."
                  className="w-full pl-8 pr-3 py-1.5 border border-slate-300 text-xs bg-white focus:outline-emerald-600 font-medium"
                />
              </div>
              {searchTerm && (
                <button
                  type="button"
                  onClick={() => setSearchTerm("")}
                  className="text-slate-500 hover:text-slate-700 text-xs font-bold px-2 py-1 bg-slate-200"
                >
                  Xóa tìm
                </button>
              )}
            </div>

            <div className="flex items-center gap-1 text-xs">
              <span className="text-slate-600 font-bold">Lọc:</span>
              <button
                type="button"
                onClick={() => setFilterStatus("ALL")}
                className={`px-2.5 py-1 font-bold border cursor-pointer ${filterStatus === "ALL" ? "bg-emerald-800 text-white border-emerald-900" : "bg-white text-slate-700 border-slate-300"}`}
              >
                Tất cả ({clusterMatrix.length})
              </button>
              <button
                type="button"
                onClick={() => setFilterStatus("HAS_OLD")}
                className={`px-2.5 py-1 font-bold border cursor-pointer ${filterStatus === "HAS_OLD" ? "bg-emerald-800 text-white border-emerald-900" : "bg-white text-slate-700 border-slate-300"}`}
              >
                Đã có xã cũ ({clusterMatrix.filter(c => c.mappedOlds.length > 0).length})
              </button>
              <button
                type="button"
                onClick={() => setFilterStatus("EMPTY")}
                className={`px-2.5 py-1 font-bold border cursor-pointer ${filterStatus === "EMPTY" ? "bg-emerald-800 text-white border-emerald-900" : "bg-white text-slate-700 border-slate-300"}`}
              >
                Chưa có xã cũ ({clusterMatrix.filter(c => c.mappedOlds.length === 0).length})
              </button>
              <button
                type="button"
                onClick={() => setFilterStatus("UNBALANCED")}
                className={`px-2.5 py-1 font-bold border cursor-pointer ${filterStatus === "UNBALANCED" ? "bg-emerald-800 text-white border-emerald-900" : "bg-white text-slate-700 border-slate-300"}`}
                title="Các xã mới có số lượng địa bàn cũ dồn về khác với số ĐBĐT mới"
              >
                Lệch số địa bàn ({clusterMatrix.filter(c => c.mappedOlds.length > 0 && c.totalOldVillages !== c.newCommune.villageCount).length})
              </button>
            </div>
          </div>
        </div>

        {/* BẢNG MA TRẬN 65 XÃ MỚI */}
        <div className="flex-1 overflow-y-auto p-3 sm:p-4">
          <div className="border border-slate-300 overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="bg-slate-100 text-slate-700 font-bold sticky top-0 border-b border-slate-300 z-10">
                <tr>
                  <th className="p-2.5 border-r border-slate-200 text-center w-12">STT</th>
                  <th className="p-2.5 border-r border-slate-200 bg-emerald-100/70 text-emerald-950 w-64">
                    XÃ / PHƯỜNG MỚI (65 XÃ GỐC)
                  </th>
                  <th className="p-2.5 border-r border-slate-200 bg-emerald-100/70 text-emerald-950 text-right w-24">
                    ĐBĐT MỚI
                  </th>
                  <th className="p-2.5 border-r border-slate-200 bg-sky-100/70 text-sky-950">
                    CÁC XÃ CŨ SÁP NHẬP VÀO XÃ NÀY (ĐÃ GOM)
                  </th>
                  <th className="p-2.5 border-r border-slate-200 bg-sky-100/70 text-sky-950 text-right w-28">
                    ĐỊA BÀN CŨ
                  </th>
                  <th className="p-2.5 border-r border-slate-200 text-center w-28">
                    CHÊNH LỆCH
                  </th>
                  <th className="p-2.5 text-center w-36">
                    THAO TÁC
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredClusters.length === 0 ? (
                  <tr>
                    <td colSpan={7} className="p-8 text-center text-slate-500 italic">
                      Không tìm thấy xã mới nào phù hợp với bộ lọc hiện tại.
                    </td>
                  </tr>
                ) : (
                  filteredClusters.map((cluster, idx) => {
                    const hasOlds = cluster.mappedOlds.length > 0;
                    const isPerfectCount = hasOlds && cluster.totalOldVillages === cluster.newCommune.villageCount;
                    const isMoreNew = hasOlds && cluster.diffVillages > 0;

                    return (
                      <tr
                        key={cluster.newCommune.name}
                        className={`hover:bg-slate-50/80 transition-colors ${
                          !hasOlds ? "bg-rose-50/30" : isPerfectCount ? "bg-emerald-50/30" : ""
                        }`}
                      >
                        {/* STT */}
                        <td className="p-2.5 border-r border-slate-200 text-center font-mono text-slate-500">
                          {idx + 1}
                        </td>

                        {/* Xã mới */}
                        <td className="p-2.5 border-r border-slate-200">
                          <div className="font-bold text-emerald-950 text-sm flex items-center gap-1.5">
                            <Building2 className="w-3.5 h-3.5 text-emerald-700 shrink-0" />
                            <span>{cluster.newCommune.name}</span>
                          </div>
                          {cluster.newCommune.code && (
                            <div className="text-[11px] font-mono text-slate-500 pl-5">
                              Mã xã: {cluster.newCommune.code}
                            </div>
                          )}
                        </td>

                        {/* Số ĐBĐT mới */}
                        <td className="p-2.5 border-r border-slate-200 text-right font-mono font-bold text-emerald-900 text-sm">
                          {cluster.newCommune.villageCount}
                        </td>

                        {/* Các xã cũ sáp nhập vào */}
                        <td className="p-2.5 border-r border-slate-200">
                          {hasOlds ? (
                            <div className="flex flex-wrap items-center gap-1.5">
                              {cluster.mappedOlds.map(old => (
                                <span
                                  key={old.name}
                                  className="inline-flex items-center gap-1.5 bg-sky-100 text-sky-950 px-2 py-1 border border-sky-300 font-bold text-xs"
                                >
                                  <span>{old.name}</span>
                                  <span className="bg-sky-200 text-sky-900 px-1 py-0.2 text-[10px] font-mono">
                                    {old.villageCount} đb
                                  </span>
                                  <button
                                    type="button"
                                    onClick={() => handleRemoveOldFromNew(old.name)}
                                    className="text-rose-600 hover:text-rose-800 hover:bg-rose-200/50 p-0.5 cursor-pointer rounded-xs border-0 bg-transparent"
                                    title={`Gỡ ${old.name} ra khỏi ${cluster.newCommune.name}`}
                                  >
                                    ✕
                                  </button>
                                </span>
                              ))}

                              {/* Nút thêm xã cũ khác vào xã mới này */}
                              {addingToNewCommune === cluster.newCommune.name ? (
                                <div className="flex items-center gap-1 bg-amber-50 p-1 border border-amber-300">
                                  <select
                                    value={selectedOldCommuneToAdd}
                                    onChange={e => setSelectedOldCommuneToAdd(e.target.value)}
                                    className="border border-slate-300 bg-white text-xs p-1 font-semibold max-w-[200px]"
                                  >
                                    <option value="">-- Chọn xã cũ để gộp --</option>
                                    {unmappedOldCommunes.map(o => (
                                      <option key={o.name} value={o.name}>
                                        {o.name} ({o.villageCount} địa bàn)
                                      </option>
                                    ))}
                                    {oldCommunesList
                                      .filter(o => !unmappedOldCommunes.some(u => u.name === o.name) && !cluster.mappedOlds.some(m => m.name === o.name))
                                      .map(o => (
                                        <option key={o.name} value={o.name}>
                                          (Chuyển từ khác): {o.name} ({o.villageCount} đb)
                                        </option>
                                      ))}
                                  </select>
                                  <button
                                    type="button"
                                    onClick={() => handleAddOldToNew(cluster.newCommune.name)}
                                    disabled={!selectedOldCommuneToAdd}
                                    className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs px-2 py-1 cursor-pointer disabled:opacity-50"
                                  >
                                    Gộp
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => setAddingToNewCommune(null)}
                                    className="text-slate-500 hover:text-slate-700 text-xs px-1.5 py-1"
                                  >
                                    Hủy
                                  </button>
                                </div>
                              ) : (
                                <button
                                  type="button"
                                  onClick={() => {
                                    setAddingToNewCommune(cluster.newCommune.name);
                                    setSelectedOldCommuneToAdd("");
                                  }}
                                  className="text-[11px] bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-2 py-1 border border-dashed border-slate-400 flex items-center gap-1 cursor-pointer"
                                  title="Thêm một xã cũ nữa gộp vào xã mới này"
                                >
                                  <Plus className="w-3 h-3 text-slate-600" />
                                  <span>Thêm xã cũ...</span>
                                </button>
                              )}
                            </div>
                          ) : (
                            <div className="flex items-center gap-2">
                              <span className="text-rose-700 italic font-semibold">
                                Chưa có xã cũ nào được gom vào xã này
                              </span>
                              {addingToNewCommune === cluster.newCommune.name ? (
                                <div className="flex items-center gap-1 bg-amber-50 p-1 border border-amber-300">
                                  <select
                                    value={selectedOldCommuneToAdd}
                                    onChange={e => setSelectedOldCommuneToAdd(e.target.value)}
                                    className="border border-slate-300 bg-white text-xs p-1 font-semibold max-w-[200px]"
                                  >
                                    <option value="">-- Chọn xã cũ --</option>
                                    {unmappedOldCommunes.map(o => (
                                      <option key={o.name} value={o.name}>
                                        {o.name} ({o.villageCount} địa bàn)
                                      </option>
                                    ))}
                                    {oldCommunesList.map(o => (
                                      <option key={o.name} value={o.name}>
                                        {o.name} ({o.villageCount} đb)
                                      </option>
                                    ))}
                                  </select>
                                  <button
                                    type="button"
                                    onClick={() => handleAddOldToNew(cluster.newCommune.name)}
                                    disabled={!selectedOldCommuneToAdd}
                                    className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs px-2 py-1 cursor-pointer disabled:opacity-50"
                                  >
                                    Gộp
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => setAddingToNewCommune(null)}
                                    className="text-slate-500 hover:text-slate-700 text-xs px-1.5 py-1"
                                  >
                                    Hủy
                                  </button>
                                </div>
                              ) : (
                                <button
                                  type="button"
                                  onClick={() => {
                                    setAddingToNewCommune(cluster.newCommune.name);
                                    setSelectedOldCommuneToAdd("");
                                  }}
                                  className="text-[11px] bg-rose-100 hover:bg-rose-200 text-rose-900 font-bold px-2 py-1 border border-rose-300 flex items-center gap-1 cursor-pointer"
                                >
                                  <Plus className="w-3 h-3 text-rose-700" />
                                  <span>Chọn xã cũ gộp vào...</span>
                                </button>
                              )}
                            </div>
                          )}
                        </td>

                        {/* Tổng địa bàn cũ dồn về */}
                        <td className="p-2.5 border-r border-slate-200 text-right font-mono font-bold text-sky-900 text-sm">
                          {cluster.totalOldVillages > 0 ? (
                            <div>
                              <span>{cluster.totalOldVillages}</span>
                              <span className="text-[10px] text-slate-500 block font-normal">
                                ({cluster.mappedOlds.length} xã cũ)
                              </span>
                            </div>
                          ) : (
                            <span className="text-slate-400">0</span>
                          )}
                        </td>

                        {/* Chênh lệch địa bàn */}
                        <td className="p-2.5 border-r border-slate-200 text-center font-bold">
                          {!hasOlds ? (
                            <span className="text-slate-400">—</span>
                          ) : isPerfectCount ? (
                            <span className="inline-flex items-center gap-1 text-emerald-700 bg-emerald-100 px-2 py-0.5 border border-emerald-300 text-[11px]">
                              <CheckCircle2 className="w-3 h-3" />
                              Khớp bằng ({cluster.newCommune.villageCount})
                            </span>
                          ) : isMoreNew ? (
                            <span className="text-sky-800 bg-sky-50 px-2 py-0.5 border border-sky-300 text-[11px] font-mono">
                              +{cluster.diffVillages} ĐBĐT mới
                            </span>
                          ) : (
                            <span className="text-amber-800 bg-amber-50 px-2 py-0.5 border border-amber-300 text-[11px] font-mono">
                              {cluster.diffVillages} địa bàn
                            </span>
                          )}
                        </td>

                        {/* Thao tác */}
                        <td className="p-2.5 text-center">
                          {hasOlds && (
                            <button
                              type="button"
                              onClick={() => {
                                // Chọn xã cũ đầu tiên của cụm để lọc trong bảng chính
                                onSelectCommuneToFilter(cluster.mappedOlds[0].name);
                                onClose();
                              }}
                              className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-[11px] px-2.5 py-1.5 flex items-center justify-center gap-1 cursor-pointer border-0 shadow-xs w-full"
                              title={`Lọc bảng làm việc chỉ cho cụm ${cluster.newCommune.name} (${cluster.mappedOlds.length} xã cũ gom về)`}
                            >
                              <Play className="w-3 h-3 fill-current" />
                              <span>Lọc &amp; Đối Soát</span>
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

        {/* FOOTER MODAL */}
        <div className="bg-slate-100 p-3 border-t border-slate-300 flex flex-wrap items-center justify-between gap-3 text-xs shrink-0">
          <div className="text-slate-600">
            💡 <strong>Mẹo nghiệp vụ:</strong> Khi gom đủ các xã cũ vào 65 xã mới, tổng số địa bàn cũ sẽ hội tụ về đủ 100%. Bấm nút <strong>"Lọc &amp; Đối Soát"</strong> ở xã nào để hệ thống lọc ngay cụm xã đó ra làm dứt điểm.
          </div>

          <button
            type="button"
            onClick={onClose}
            className="bg-slate-700 hover:bg-slate-800 text-white font-bold px-4 py-1.5 cursor-pointer border-0"
          >
            Đóng Lại &amp; Tiếp Tục Đối Soát
          </button>
        </div>
      </div>
    </div>
  );
};
