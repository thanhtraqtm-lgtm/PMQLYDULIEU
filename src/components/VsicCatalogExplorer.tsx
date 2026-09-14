import React, { useState, useMemo } from "react";
import { Search, BookOpen, Filter, Download, Copy, Check, Sparkles, ChevronRight } from "lucide-react";
import * as XLSX from "xlsx";
import { 
  vsicRawData, 
  getSectorLevel, 
  smartSuggestSectorByDescription, 
  getParentSectorCode 
} from "../data/vsic";

export const VsicCatalogExplorer: React.FC = () => {
  const [searchTerm, setSearchTerm] = useState("");
  const [levelFilter, setLevelFilter] = useState<number | 0>(0); // 0 = all
  const [copiedCode, setCopiedCode] = useState<string | null>(null);

  // Smart suggestion tool state
  const [descInput, setDescInput] = useState("");
  const [suggestedResults, setSuggestedResults] = useState<Array<{ code: string; name: string; score?: number }>>([]);

  // Transform vsicRawData into array
  const allSectors = useMemo(() => {
    return Object.entries(vsicRawData).map(([code, name]) => {
      const level = getSectorLevel(code);
      const parent = getParentSectorCode(code);
      return { code, name, level, parent };
    });
  }, []);

  // Filtered sectors based on search & level
  const filteredSectors = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    return allSectors.filter(item => {
      if (levelFilter !== 0 && item.level !== levelFilter) {
        return false;
      }
      if (!term) return true;
      return (
        item.code.toLowerCase().includes(term) ||
        item.name.toLowerCase().includes(term)
      );
    });
  }, [allSectors, searchTerm, levelFilter]);

  // Copy code to clipboard
  const handleCopy = (code: string) => {
    navigator.clipboard.writeText(code);
    setCopiedCode(code);
    setTimeout(() => setCopiedCode(null), 1500);
  };

  // Export VSIC to Excel
  const handleExport = () => {
    const dataToExport = filteredSectors.map((s, idx) => ({
      "STT": idx + 1,
      "Cấp": s.level,
      "Mã ngành VSIC": s.code,
      "Tên ngành kinh tế": s.name,
      "Mã cấp cha": s.parent || ""
    }));

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DanhMuc_VSIC_2018");
    XLSX.writeFile(wb, `DanhMuc_VSIC_QĐ27_2018_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Handle smart suggestion
  const handleSmartSuggest = () => {
    if (!descInput.trim()) return;
    const topResult = smartSuggestSectorByDescription(descInput);
    if (topResult) {
      const matched = allSectors.filter(s => 
        s.code === topResult.ma || 
        s.code.startsWith(topResult.ma) || 
        topResult.ma.startsWith(s.code)
      );
      setSuggestedResults(
        matched.length > 0 
          ? matched 
          : [{ code: topResult.ma, name: topResult.ten }]
      );
    } else {
      setSuggestedResults([]);
    }
  };

  return (
    <div className="bg-sky-50/30 border border-sky-200 rounded-none shadow-xs overflow-hidden font-sans space-y-3">
      {/* THANH TIÊU ĐỀ CHÍNH - Đồng bộ phong cách #286e42 */}
      <div className="px-3.5 py-2.5 bg-sky-100/60 border-b border-sky-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex items-center gap-2.5 min-w-0">
          <div className="w-7 h-7 bg-[#286e42] flex items-center justify-center text-white shrink-0 rounded-none shadow-2xs">
            <BookOpen className="w-4 h-4" />
          </div>
          <div className="min-w-0">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-bold text-slate-900 uppercase tracking-wide">
                Hệ Thống Ngành Kinh Tế Việt Nam (VSIC 2018)
              </span>
              <span className="text-[11px] font-mono text-sky-900 bg-sky-200/80 px-2 py-0.5 rounded-none font-bold border border-sky-300">
                {filteredSectors.length} ngành • QĐ 27/2018/QĐ-TTg
              </span>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={handleExport}
            className="bg-[#286e42] hover:bg-[#205835] text-white font-bold text-xs px-3 py-1 rounded-none transition-colors shadow-2xs flex items-center gap-1 cursor-pointer whitespace-nowrap border-0"
            title="Xuất danh mục ngành kinh tế ra file Excel"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Xuất Excel ({filteredSectors.length})</span>
          </button>
        </div>
      </div>

      {/* THANH GỢI Ý MÃ NGÀNH THÔNG MINH */}
      <div className="mx-3.5 bg-white border border-sky-200 rounded-none p-3 space-y-2.5">
        <div className="flex items-center gap-2">
          <Sparkles className="w-4 h-4 text-[#0284c7]" />
          <span className="text-xs font-bold text-slate-800">
            Gợi ý mã ngành thông minh từ mô tả hoạt động kinh doanh:
          </span>
        </div>
        <div className="flex flex-col sm:flex-row gap-2">
          <input
            type="text"
            value={descInput}
            onChange={(e) => setDescInput(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleSmartSuggest()}
            placeholder="Nhập mô tả hoạt động (ví dụ: bán buôn máy vi tính, may trang phục, trồng cây ăn quả...)"
            className="flex-1 px-3 py-1.5 text-xs bg-white border border-slate-300 rounded-none focus:outline-none focus:border-sky-500 text-slate-800 placeholder-slate-400 font-medium"
          />
          <button
            type="button"
            onClick={handleSmartSuggest}
            className="px-3.5 py-1.5 text-xs font-bold text-white bg-[#0284c7] hover:bg-[#0369a1] rounded-none transition shadow-2xs flex items-center justify-center gap-1.5 cursor-pointer shrink-0 border-0"
          >
            <Sparkles className="w-3.5 h-3.5" />
            <span>Gợi ý mã VSIC</span>
          </button>
        </div>

        {suggestedResults.length > 0 && (
          <div className="pt-2 border-t border-sky-100">
            <span className="text-[11px] font-bold text-sky-900 block mb-1.5">Kết quả đề xuất phù hợp nhất:</span>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
              {suggestedResults.slice(0, 6).map((item) => (
                <div 
                  key={item.code} 
                  className="p-2 bg-sky-50/50 rounded-none border border-sky-200 flex items-start justify-between gap-2"
                >
                  <div>
                    <div className="flex items-center gap-1.5">
                      <span className="font-mono font-bold text-xs text-sky-800 bg-white px-1.5 py-0.5 border border-sky-300">
                        {item.code}
                      </span>
                      <span className="text-[10px] text-slate-500 font-medium">Cấp {getSectorLevel(item.code)}</span>
                    </div>
                    <p className="text-xs text-slate-800 mt-1 line-clamp-2 leading-tight font-medium">{item.name}</p>
                  </div>
                  <button
                    type="button"
                    onClick={() => handleCopy(item.code)}
                    title="Sao chép mã"
                    className="p-1 hover:bg-white rounded-none text-slate-400 hover:text-slate-700 cursor-pointer shrink-0 border border-transparent hover:border-slate-200"
                  >
                    {copiedCode === item.code ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* BỘ LỌC TÌM KIẾM & CẤP NGÀNH */}
      <div className="mx-3.5 bg-white border border-sky-200 rounded-none p-3 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
        <div className="relative flex-1">
          <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Tìm kiếm theo mã số (01, 10, C, 4711) hoặc từ khóa tên ngành..."
            className="w-full pl-8 pr-8 py-1.5 text-xs bg-white border border-slate-300 rounded-none focus:outline-none focus:border-sky-500 text-slate-800 placeholder-slate-400 font-medium"
          />
          {searchTerm && (
            <button
              type="button"
              onClick={() => setSearchTerm("")}
              className="absolute right-2.5 top-1/2 -translate-y-1/2 text-xs text-slate-400 hover:text-slate-600 bg-transparent border-0 cursor-pointer font-bold"
            >
              ✕
            </button>
          )}
        </div>

        {/* Nút lọc cấp */}
        <div className="flex items-center gap-1 overflow-x-auto">
          <span className="text-xs font-bold text-slate-600 mr-1 flex items-center gap-1 shrink-0">
            <Filter className="w-3.5 h-3.5" /> Cấp:
          </span>
          {[
            { level: 0, label: "Tất cả" },
            { level: 1, label: "Cấp 1 (A-U)" },
            { level: 2, label: "Cấp 2 (2 số)" },
            { level: 3, label: "Cấp 3 (3 số)" },
            { level: 4, label: "Cấp 4 (4 số)" },
            { level: 5, label: "Cấp 5 (5 số)" },
          ].map((f) => (
            <button
              key={f.level}
              type="button"
              onClick={() => setLevelFilter(f.level)}
              className={`px-2.5 py-1 text-xs font-bold rounded-none border transition whitespace-nowrap cursor-pointer ${
                levelFilter === f.level
                  ? "bg-[#286e42] text-white border-[#1d4f2f] shadow-2xs"
                  : "bg-slate-50 text-slate-700 border-slate-300 hover:bg-slate-100"
              }`}
            >
              {f.label}
            </button>
          ))}
        </div>
      </div>

      {/* BẢNG DANH MỤC VSIC VỚI HEADER #286e42 */}
      <div className="mx-3.5 border border-sky-200 rounded-none overflow-hidden max-h-[520px] overflow-y-auto bg-white shadow-inner">
        <table className="w-full text-left border-collapse text-xs">
          <thead className="bg-[#286e42] sticky top-0 text-white font-bold border-b border-[#1d4f2f] z-10">
            <tr>
              <th className="py-2 px-3 w-14 text-center border-r border-[#1d4f2f]">STT</th>
              <th className="py-2 px-3 w-20 text-center border-r border-[#1d4f2f]">CẤP</th>
              <th className="py-2 px-3 w-28 border-r border-[#1d4f2f]">MÃ NGÀNH</th>
              <th className="py-2 px-3 border-r border-[#1d4f2f]">TÊN NGÀNH KINH TẾ VIỆT NAM</th>
              <th className="py-2 px-3 w-24 text-center border-r border-[#1d4f2f]">MÃ CHA</th>
              <th className="py-2 px-3 w-16 text-center">CHÉP</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filteredSectors.length === 0 ? (
              <tr>
                <td colSpan={6} className="py-8 text-center text-slate-400 italic">
                  Không tìm thấy mã hoặc tên ngành nào phù hợp với bộ lọc hiện tại.
                </td>
              </tr>
            ) : (
              filteredSectors.map((sector, index) => {
                const isLevel1 = sector.level === 1;
                const isLevel2 = sector.level === 2;

                return (
                  <tr
                    key={sector.code}
                    className={`hover:bg-sky-50/50 ${
                      isLevel1
                        ? "bg-emerald-50/50 font-bold text-slate-900"
                        : isLevel2
                        ? "bg-slate-50/60 font-semibold text-slate-800"
                        : index % 2 === 1
                        ? "bg-slate-50/30"
                        : ""
                    }`}
                  >
                    <td className="py-1.5 px-3 text-center text-slate-400 font-mono text-[11px] border-r border-slate-100">
                      {index + 1}
                    </td>
                    <td className="py-1.5 px-3 text-center border-r border-slate-100">
                      <span
                        className={`inline-block px-1.5 py-0.5 text-[10px] font-bold border rounded-none ${
                          sector.level === 1
                            ? "bg-purple-100 text-purple-800 border-purple-300"
                            : sector.level === 2
                            ? "bg-blue-100 text-blue-800 border-blue-300"
                            : sector.level === 3
                            ? "bg-emerald-100 text-emerald-800 border-emerald-300"
                            : sector.level === 4
                            ? "bg-amber-100 text-amber-800 border-amber-300"
                            : "bg-slate-100 text-slate-700 border-slate-300"
                        }`}
                      >
                        Cấp {sector.level}
                      </span>
                    </td>
                    <td className="py-1.5 px-3 font-mono font-bold text-sky-800 border-r border-slate-100">
                      {sector.code}
                    </td>
                    <td className="py-1.5 px-3 border-r border-slate-100">
                      <div 
                        className="flex items-center gap-1.5"
                        style={{ paddingLeft: `${(sector.level - 1) * 12}px` }}
                      >
                        {sector.level > 1 && (
                          <ChevronRight className="w-3 h-3 text-slate-400 shrink-0" />
                        )}
                        <span className="font-medium">{sector.name}</span>
                      </div>
                    </td>
                    <td className="py-1.5 px-3 text-center font-mono text-slate-500 border-r border-slate-100">
                      {sector.parent || "—"}
                    </td>
                    <td className="py-1.5 px-3 text-center">
                      <button
                        type="button"
                        onClick={() => handleCopy(sector.code)}
                        title="Sao chép mã ngành"
                        className="p-1 hover:bg-slate-200 text-slate-400 hover:text-slate-800 rounded-none transition cursor-pointer border-0 bg-transparent"
                      >
                        {copiedCode === sector.code ? (
                          <Check className="w-3.5 h-3.5 text-emerald-600" />
                        ) : (
                          <Copy className="w-3.5 h-3.5" />
                        )}
                      </button>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default VsicCatalogExplorer;
