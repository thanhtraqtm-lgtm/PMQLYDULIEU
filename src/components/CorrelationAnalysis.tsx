import React, { useState, useMemo, useEffect } from "react";
import { Activity, Search, AlertTriangle } from "lucide-react";
import * as XLSX from "xlsx";

// Helper to calculate normal cumulative distribution function (CDF)
function normalCDF(x: number): number {
  const t = 1 / (1 + 0.2316419 * Math.abs(x));
  const d = 0.3989422804; // 1 / Math.sqrt(2 * Math.PI)
  const p = 1 - d * Math.exp(-x * x / 2) * (
    0.31938153 * t -
    0.356563782 * t * t +
    1.781477937 * Math.pow(t, 3) -
    1.821255978 * Math.pow(t, 4) +
    1.330274429 * Math.pow(t, 5)
  );
  return x >= 0 ? p : 1 - p;
}

// Helper to calculate the p-value of a Chi-Square statistic
function chiSquarePValue(chiSq: number, df: number): number {
  if (df <= 0) return 1.0;
  if (chiSq <= 0) return 1.0;

  if (df === 1) {
    return 2 * (1 - normalCDF(Math.sqrt(chiSq)));
  } else if (df === 2) {
    return Math.exp(-chiSq / 2);
  } else {
    // Wilson-Hilferty transformation of Chi-Square to normal distribution
    const z = (Math.pow(chiSq / df, 1/3) - (1 - 2/(9*df))) / Math.sqrt(2/(9*df));
    return 1 - normalCDF(z);
  }
}

interface CorrelationAnalysisProps {
  mainData: any[];
  columns: string[];
  setRowIndicesFilter: (indices: number[]) => void;
  setRowFilterLabel: (label: string) => void;
  setViewPage: (page: number) => void;
  setActiveTab: (tab: string) => void;
}

export const CorrelationAnalysis = React.memo(function CorrelationAnalysis({
  mainData,
  columns,
  setRowIndicesFilter,
  setRowFilterLabel,
  setViewPage,
  setActiveTab
}: CorrelationAnalysisProps) {
  // States for sub-tab selection and calculation triggers
  const [tqSubTab, setTqSubTab] = useState<"bang_cheo" | "tuyen_tinh">("bang_cheo");
  const [isTqCalculated, setIsTqCalculated] = useState<boolean>(false);

  // States for Bảng chéo (Crosstab)
  const [tqHangCol, setTqHangCol] = useState<string>("");
  const [tqCotCol, setTqCotCol] = useState<string>("");
  const [tqShowResults, setTqShowResults] = useState<boolean>(false);
  const [tqSearchTerm, setTqSearchTerm] = useState<string>("");

  // States for detailed cell household modal
  const [tqSelectedCell, setTqSelectedCell] = useState<{ hang: string | null; cot: string | null } | null>(null);
  const [tqModalSearchTerm, setTqModalSearchTerm] = useState<string>("");
  const [tqModalPage, setTqModalPage] = useState<number>(1);

  // States for Tương quan tuyến tính (Linear/Comparison)
  const [tqSelectedCol1, setTqSelectedCol1] = useState<string>("");
  const [tqSelectedCol2, setTqSelectedCol2] = useState<string>("");

  // Auto-detect default row/column variables for crosstab without overwriting active user selection
  useEffect(() => {
    if (columns.length > 0) {
      setTqHangCol(prev => {
        if (prev && columns.includes(prev)) return prev;
        return columns.find(c => c.toLowerCase().includes("ngành") || c.toLowerCase().includes("nghề") || c.toLowerCase().includes("mã") || c.toLowerCase().includes("xã")) || columns[0];
      });
      
      setTqCotCol(prev => {
        if (prev && columns.includes(prev)) return prev;
        const currentHang = tqHangCol || columns[0];
        return columns.find(c => (c.toLowerCase().includes("cây") || c.toLowerCase().includes("trồng") || c.toLowerCase().includes("thu")) && c !== currentHang) || (columns[1] || columns[0]);
      });
    }
  }, [columns]);

  // Compute Crosstab and Chi-Square
  const crosstabResult = useMemo(() => {
    if (!mainData || mainData.length === 0) return null;
    const col1 = tqHangCol || (columns.length > 0 ? columns[0] : "");
    const col2 = tqCotCol || (columns.length > 1 ? columns[1] : (columns[0] || ""));

    if (!col1 || !col2) return null;

    const rawRowsSet = new Set<string>();
    const rawColsSet = new Set<string>();
    const matrix: Record<string, Record<string, number>> = {};
    const rowTotals: Record<string, number> = {};
    const colTotals: Record<string, number> = {};
    let grandTotal = 0;

    mainData.forEach(row => {
      let rVal = String(row[col1] ?? "").trim();
      let cVal = String(row[col2] ?? "").trim();
      if (!rVal) rVal = "(Trống)";
      if (!cVal) cVal = "(Trống)";

      rawRowsSet.add(rVal);
      rawColsSet.add(cVal);

      if (!matrix[rVal]) matrix[rVal] = {};
      matrix[rVal][cVal] = (matrix[rVal][cVal] || 0) + 1;

      rowTotals[rVal] = (rowTotals[rVal] || 0) + 1;
      colTotals[cVal] = (colTotals[cVal] || 0) + 1;
      grandTotal++;
    });

    const smartSort = (arr: string[]) => {
      return [...arr].sort((a, b) => {
        const aLow = a.toLowerCase().trim();
        const bLow = b.toLowerCase().trim();
        if (aLow === "có") return -1;
        if (bLow === "có") return 1;
        if (aLow === "không") {
          if (bLow === "có") return 1;
          return -1;
        }
        if (bLow === "không") {
          if (aLow === "có") return -1;
          return 1;
        }
        return a.localeCompare(b, "vi");
      });
    };

    const sortedRows = smartSort(Array.from(rawRowsSet));
    const sortedCols = smartSort(Array.from(rawColsSet));

    // Calculate Chi-Square test of Independence
    let chiSquare = 0;
    let df = 0;
    let pValue = 1.0;
    let chiSquareValid = false;

    const rCount = sortedRows.length;
    const cCount = sortedCols.length;

    if (rCount > 1 && cCount > 1 && grandTotal > 0) {
      chiSquareValid = true;
      df = (rCount - 1) * (cCount - 1);

      sortedRows.forEach(rVal => {
        const rTot = rowTotals[rVal] || 0;
        sortedCols.forEach(cVal => {
          const cTot = colTotals[cVal] || 0;
          const expected = (rTot * cTot) / grandTotal;
          if (expected > 0) {
            const observed = matrix[rVal]?.[cVal] || 0;
            chiSquare += Math.pow(observed - expected, 2) / expected;
          }
        });
      });

      pValue = chiSquarePValue(chiSquare, df);
    }

    return {
      col1,
      col2,
      matrix,
      rowTotals,
      colTotals,
      grandTotal,
      sortedRows,
      sortedCols,
      chiSquare,
      df,
      pValue,
      chiSquareValid
    };
  }, [mainData, tqHangCol, tqCotCol, columns]);

  if (mainData.length === 0) {
    return (
      <div className="bg-slate-50 rounded-2xl p-12 text-center text-xs text-amber-700 border border-slate-200 font-condensed space-y-4 font-sans">
        <div className="flex justify-center">
          <AlertTriangle className="w-8 h-8 text-amber-500 animate-bounce" />
        </div>
        <div>
          <h3 className="text-sm font-bold uppercase tracking-wider text-slate-800">CHƯA CÓ DỮ LIỆU NGUỒN KHẢO SÁT</h3>
          <p className="text-xs text-slate-500 max-w-md mx-auto mt-1 font-sans">
            Hãy nạp file Excel hoặc CSV danh sách đơn vị cơ sở khảo sát tại Trang Chủ trước để tiến hành phân tích tương quan.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6 animate-fade-in font-sans">
      {/* Tab Header Banner */}
      <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-slate-100 pb-4">
          <div className="space-y-1">
            <h3 className="text-base font-extrabold uppercase tracking-wider text-slate-800 flex items-center gap-2">
              <Activity className="w-5 h-5 text-emerald-600 animate-pulse" />
              PHÂN TÍCH TƯƠNG QUAN HỘ &amp; CHỈ TIÊU LIÊN KẾT
            </h3>
            <p className="text-xs text-slate-500">
              Đo lường mối tương quan liên hệ giữa các đặc trưng, xây dựng bảng liên đới chéo (Cross-Tabulation) trực quan hoặc đo hệ số Pearson tương quan tuyến tính.
            </p>
          </div>

          <div className="flex border border-slate-200 rounded-xl overflow-hidden shrink-0 bg-slate-50 p-1">
            <button
              onClick={() => setTqSubTab("bang_cheo")}
              className={`px-4 py-2 text-xs font-bold transition-all cursor-pointer rounded-lg border-0 flex items-center gap-1.5 ${
                tqSubTab === "bang_cheo"
                  ? "bg-emerald-600 text-white shadow-sm font-black"
                  : "bg-transparent text-slate-600 hover:text-slate-900"
              }`}
            >
              📊 Thống kê tương quan hộ (Bảng Chéo)
            </button>
            <button
              onClick={() => setTqSubTab("tuyen_tinh")}
              className={`px-4 py-2 text-xs font-bold transition-all cursor-pointer rounded-lg border-0 flex items-center gap-1.5 ${
                tqSubTab === "tuyen_tinh"
                  ? "bg-indigo-600 text-white shadow-sm font-black"
                  : "bg-transparent text-slate-600 hover:text-slate-900"
              }`}
            >
              📈 Tương quan tuyến tính &amp; Trung bình
            </button>
          </div>
        </div>

        {!isTqCalculated ? (
          <div className="bg-slate-900 text-slate-100 rounded-2xl p-8 border border-slate-800 flex flex-col items-center justify-center text-center space-y-4 max-w-xl mx-auto my-6 shadow-md animate-fade-in font-sans">
            <div className="bg-emerald-500/10 text-emerald-400 p-4 rounded-full border border-emerald-500/20">
              <Activity className="w-8 h-8 text-emerald-400 animate-pulse" />
            </div>
            <div className="space-y-1.5">
              <h4 className="text-sm font-black tracking-wider uppercase text-slate-200">
                SẴN SÀNG PHÂN TÍCH TƯƠNG QUAN &amp; BẢNG CHÉO
              </h4>
              <p className="text-xs text-slate-400 max-w-sm">
                Xây dựng bảng liên đới chéo (Crosstab) và thực hiện phép kiểm định độc lập Chi-Square để dò tìm các quy luật ràng buộc kinh tế - xã hội quan trọng.
              </p>
            </div>
            <button
              onClick={() => setIsTqCalculated(true)}
              className="bg-emerald-600 hover:bg-emerald-550 text-white font-bold text-xs py-2.5 px-6 rounded-xl border-0 cursor-pointer shadow transition-all active:scale-95 flex items-center gap-1.5"
            >
              ⚡ KHỞI CHẠY PHÂN TÍCH TƯƠNG QUAN CHỈ SỐ
            </button>
          </div>
        ) : (
          tqSubTab === "bang_cheo" ? (
            // ================= SUB TAB 1: BẢNG CHÉO LIÊN ĐỚI TƯƠNG QUAN HỘ =================
            (() => {
              if (!crosstabResult) {
                return (
                  <div className="p-6 text-center text-slate-500 italic">
                    Đang xử lý dữ liệu tương quan...
                  </div>
                );
              }

              const {
                col1,
                col2,
                matrix,
                rowTotals,
                colTotals,
                grandTotal,
                sortedRows,
                sortedCols,
                chiSquare,
                df,
                pValue,
                chiSquareValid
              } = crosstabResult;

              const displayedRows = tqSearchTerm.trim()
                ? sortedRows.filter(r => r.toLowerCase().includes(tqSearchTerm.toLowerCase()))
                : sortedRows;

              const handleExportCrosstab = () => {
                try {
                  const tableRows: (string | number)[][] = [];
                  const header = [`Chỉ tiêu hàng: ${col1} \\ Chỉ tiêu cột: ${col2}`, ...sortedCols, "Tổng cộng"];
                  tableRows.push(header);
                  
                  const colTotalsRow: (string | number)[] = ["Tổng cộng"];
                  sortedCols.forEach(cVal => {
                    colTotalsRow.push(colTotals[cVal] || 0);
                  });
                  colTotalsRow.push(grandTotal);
                  tableRows.push(colTotalsRow);
                  
                  sortedRows.forEach(rVal => {
                    const rowArr: (string | number)[] = [rVal];
                    sortedCols.forEach(cVal => {
                      rowArr.push(matrix[rVal]?.[cVal] || 0);
                    });
                    rowArr.push(rowTotals[rVal] || 0);
                    tableRows.push(rowArr);
                  });

                  tableRows.push([]);
                  tableRows.push(["KẾT QUẢ KIỂM ĐỊNH CHI-SQUARE (INDEPENDENCE TEST)"]);
                  tableRows.push(["Giá trị Chi-Square", chiSquareValid ? chiSquare.toFixed(4) : "N/A"]);
                  tableRows.push(["Bậc tự do (df)", chiSquareValid ? df : "N/A"]);
                  tableRows.push(["Giá trị p-value", chiSquareValid ? pValue.toFixed(6) : "N/A"]);
                  tableRows.push([
                    "Ý nghĩa thống kê",
                    chiSquareValid
                      ? pValue < 0.05
                        ? "Có ý nghĩa thống kê ở mức ý nghĩa 5% (Hai chỉ tiêu có tương quan liên hệ mật thiết)"
                        : "Không có ý nghĩa thống kê ở mức ý nghĩa 5% (Hai chỉ tiêu độc lập thống kê)"
                      : "Không đủ dữ liệu để kiểm định"
                  ]);

                  const ws = XLSX.utils.aoa_to_sheet(tableRows);
                  const wb = XLSX.utils.book_new();
                  XLSX.utils.book_append_sheet(wb, ws, "Bảng chéo tương quan");
                  XLSX.writeFile(wb, `Thong_Ke_Tuong_Quan_Ho_${new Date().toISOString().slice(0, 10)}.xlsx`);
                } catch (err: any) {
                  alert("Lỗi xuất Excel bảng chéo: " + err.message);
                }
              };

              const handleFilterHouseholds = (hangVal: string | null, cotVal: string | null) => {
                setTqSelectedCell({ hang: hangVal, cot: cotVal });
                setTqModalSearchTerm("");
                setTqModalPage(1);
              };

              return (
                <div className="space-y-5">
                  <div className="bg-gradient-to-r from-emerald-50/20 via-slate-50 to-slate-50 border border-slate-200 rounded-2xl p-4 space-y-4">
                    <div className="flex flex-col lg:flex-row lg:items-end justify-between gap-4">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 flex-1">
                        <div className="space-y-1.5">
                          <label className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider font-mono">
                            Chỉ tiêu hàng
                          </label>
                          <select
                            value={col1}
                            onChange={(e) => {
                              setTqHangCol(e.target.value);
                              setIsTqCalculated(false);
                            }}
                            className="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500 shadow-sm cursor-pointer"
                          >
                            {columns.map((col) => (
                              <option key={col} value={col}>{col}</option>
                            ))}
                          </select>
                        </div>

                        <div className="space-y-1.5">
                          <label className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider font-mono">
                            Chỉ tiêu cột
                          </label>
                          <select
                            value={col2}
                            onChange={(e) => {
                              setTqCotCol(e.target.value);
                              setIsTqCalculated(false);
                            }}
                            className="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500 shadow-sm cursor-pointer"
                          >
                            {columns.map((col) => (
                              <option key={col} value={col}>{col}</option>
                            ))}
                          </select>
                        </div>
                      </div>

                      <div className="flex flex-wrap items-center gap-2 md:self-end">
                        <button
                          onClick={() => setTqShowResults(true)}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black px-5 py-3 rounded-xl transition-all shadow-md flex items-center gap-2 cursor-pointer border-0 uppercase tracking-wider active:scale-95 shrink-0 font-sans"
                        >
                          ⚡ Xem KQ
                        </button>

                        <button
                          onClick={handleExportCrosstab}
                          className="bg-[#5cb85c] hover:bg-[#4cae4c] text-white text-xs font-black px-5 py-3 rounded-xl transition-all shadow-sm flex items-center gap-1.5 cursor-pointer border-0 active:scale-95 shrink-0 font-sans"
                        >
                          📥 Tải dữ liệu Excel
                        </button>

                        <div className="relative min-w-[180px] shrink-0 font-sans">
                          <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                          <input
                            type="text"
                            value={tqSearchTerm}
                            onChange={(e) => setTqSearchTerm(e.target.value)}
                            placeholder="Tìm dòng..."
                            className="w-full bg-white hover:border-slate-400 border border-slate-300 text-xs rounded-xl pl-9 pr-8 py-2.5 focus:outline-none focus:border-emerald-500 transition-all font-medium shadow-sm"
                          />
                          {tqSearchTerm && (
                            <button
                              onClick={() => setTqSearchTerm("")}
                              className="absolute right-3 top-1/2 -translate-y-1/2 p-0.5 text-slate-400 hover:text-slate-600 text-xs font-bold cursor-pointer bg-transparent border-0"
                            >
                              ×
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>

                  {tqShowResults && (
                    <div className="space-y-6">
                      <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                        <div className="p-4 border-b border-slate-100 bg-[#fbf9f4] flex justify-between items-center">
                          <h4 className="text-xs font-black tracking-wider text-slate-800 uppercase font-mono flex items-center gap-2">
                            <span>📊</span> Bảng Chéo Tương Quan Liên Hệ Đa Chỉ Tiêu
                          </h4>
                          <span className="text-[10px] bg-emerald-50 text-emerald-700 border border-emerald-200 px-2.5 py-1 rounded font-mono font-extrabold uppercase">
                            Tổng số: {grandTotal.toLocaleString()} hộ
                          </span>
                        </div>

                        <div className="overflow-x-auto custom-scrollbar">
                          <table className="w-full border-collapse text-xs text-center border border-slate-200">
                            <thead>
                              <tr className="bg-[#faf6eb] border-b border-slate-250 text-slate-700 font-bold">
                                <th className="p-3.5 border-r border-slate-200 text-left min-w-[200px] font-sans text-[11px] uppercase text-slate-500 tracking-wider">
                                  Chỉ tiêu: {col1}
                                </th>
                                {sortedCols.map(colVal => (
                                  <th key={colVal} className="p-3.5 border-r border-slate-200 min-w-[120px] text-center">
                                    {colVal}
                                  </th>
                                ))}
                                <th className="p-3.5 border-r border-slate-200 min-w-[120px] text-center bg-slate-50/80 text-slate-800 font-extrabold">
                                  Tổng cộng
                                </th>
                              </tr>
                            </thead>
                            <tbody>
                              <tr className="bg-[#faf6eb]/40 border-b border-slate-200 font-extrabold hover:bg-slate-50/50 transition-colors">
                                <td className="p-3.5 border-r border-slate-200 text-left font-sans text-slate-700">
                                  Tổng cộng
                                </td>
                                {sortedCols.map(colVal => {
                                  const count = colTotals[colVal] || 0;
                                  return (
                                    <td key={colVal} className="p-3.5 border-r border-slate-200 text-center">
                                      {count > 0 ? (
                                        <button
                                          onClick={() => handleFilterHouseholds(null, colVal)}
                                          className="font-mono text-xs text-sky-600 hover:text-sky-800 hover:underline font-bold bg-transparent border-0 cursor-pointer p-0"
                                          title={`Click để xem danh sách các hộ có ${col2} = ${colVal}`}
                                        >
                                          {count.toLocaleString()}
                                        </button>
                                      ) : (
                                        <span className="text-slate-400 font-mono">0</span>
                                      )}
                                    </td>
                                  );
                                })}
                                <td className="p-3.5 border-r border-slate-200 text-center bg-[#faf6eb]/60 font-mono text-slate-900 font-bold">
                                  <button
                                    onClick={() => handleFilterHouseholds(null, null)}
                                    className="font-mono text-xs text-sky-600 hover:text-sky-800 hover:underline font-bold bg-transparent border-0 cursor-pointer p-0"
                                    title="Click để xem toàn bộ danh sách hộ"
                                  >
                                    {grandTotal.toLocaleString()}
                                  </button>
                                </td>
                              </tr>

                              {displayedRows.map(rVal => (
                                <tr key={rVal} className="border-b border-slate-100 hover:bg-slate-50/80 transition-colors">
                                  <td className="p-3.5 border-r border-slate-200 text-left font-bold text-slate-700">
                                    {rVal}
                                  </td>
                                  {sortedCols.map(colVal => {
                                    const count = matrix[rVal]?.[colVal] || 0;
                                    return (
                                      <td key={colVal} className="p-3.5 border-r border-slate-200 text-center font-mono">
                                        {count > 0 ? (
                                          <button
                                            onClick={() => handleFilterHouseholds(rVal, colVal)}
                                            className="font-mono text-xs text-sky-600 hover:text-sky-800 hover:underline font-bold bg-transparent border-0 cursor-pointer p-0"
                                            title={`Click để xem danh sách hộ: [${rVal}] × [${colVal}]`}
                                          >
                                            {count.toLocaleString()}
                                          </button>
                                        ) : (
                                          <span className="text-slate-300">0</span>
                                        )}
                                      </td>
                                    );
                                  })}
                                  <td className="p-3.5 border-r border-slate-200 text-center bg-slate-50/50 font-mono font-bold text-slate-800">
                                    {rowTotals[rVal] > 0 ? (
                                      <button
                                        onClick={() => handleFilterHouseholds(rVal, null)}
                                        className="font-mono text-xs text-sky-600 hover:text-sky-800 hover:underline font-bold bg-transparent border-0 cursor-pointer p-0"
                                        title={`Click để xem danh sách hộ có ${col1} = ${rVal}`}
                                      >
                                        {(rowTotals[rVal] || 0).toLocaleString()}
                                      </button>
                                    ) : (
                                      <span className="text-slate-400">0</span>
                                    )}
                                  </td>
                                </tr>
                              ))}

                              {displayedRows.length === 0 && (
                                <tr>
                                  <td colSpan={sortedCols.length + 2} className="p-10 text-slate-400 italic">
                                    Không tìm thấy kết quả phù hợp với từ khóa tìm kiếm.
                                  </td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                      </div>

                      <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm space-y-4">
                        <div className="flex items-center gap-2 pb-3 border-b border-slate-100">
                          <div className="p-2 bg-indigo-50 text-indigo-600 rounded-xl">
                            <Activity className="w-5 h-5 animate-pulse" />
                          </div>
                          <div>
                            <h4 className="text-xs font-black uppercase tracking-wider text-slate-800 font-sans">
                              PHÉP KIỂM ĐỊNH ĐỘC LẬP CHI-SQUARE (INDEPENDENCE TEST)
                            </h4>
                            <p className="text-[11px] text-slate-550 font-sans">
                              Đánh giá xem liệu hai chỉ tiêu "{col1}" và "{col2}" độc lập thống kê hay có mối tương quan phụ thuộc lẫn nhau.
                            </p>
                          </div>
                        </div>

                        {chiSquareValid ? (
                          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 pt-1">
                            <div className="lg:col-span-1 bg-slate-50/80 border border-slate-200 rounded-2xl p-4 flex flex-col justify-center space-y-3 font-sans">
                              <div className="flex items-center justify-between border-b border-slate-200/60 pb-2">
                                <span className="text-[11px] font-bold text-slate-500 font-mono uppercase">Trị số Chi-Square (χ²)</span>
                                <span className="font-mono font-bold text-sm text-slate-800">{chiSquare.toFixed(4)}</span>
                              </div>
                              <div className="flex items-center justify-between border-b border-slate-200/60 pb-2">
                                <span className="text-[11px] font-bold text-slate-500 font-mono uppercase">Bậc tự do (df)</span>
                                <span className="font-mono font-bold text-sm text-slate-800">{df}</span>
                              </div>
                              <div className="flex items-center justify-between">
                                <span className="text-[11px] font-bold text-slate-500 font-mono uppercase">Trị số p-value (p)</span>
                                <span className="font-mono font-extrabold text-sm text-indigo-700">
                                  {pValue < 0.0001 ? "< 0.0001" : pValue.toFixed(4)}
                                </span>
                              </div>
                            </div>

                            <div className="lg:col-span-2 flex flex-col justify-between font-sans">
                              <div className={`border rounded-2xl p-4.5 flex-1 flex flex-col justify-between space-y-3 ${
                                pValue < 0.05 
                                  ? "bg-emerald-50/50 border-emerald-200/80 text-emerald-850" 
                                  : "bg-slate-50/80 border-slate-200 text-slate-800"
                              }`}>
                                <div className="space-y-1.5">
                                  <div className="flex items-center gap-1.5">
                                    <span className={`w-2 h-2 rounded-full ${pValue < 0.05 ? "bg-emerald-500" : "bg-slate-500"}`} />
                                    <h5 className="text-xs font-extrabold uppercase tracking-wide">
                                      Kết luận kiểm định độc lập (mức ý nghĩa 5%)
                                    </h5>
                                  </div>
                                  <div className="text-xs font-sans leading-relaxed">
                                    {pValue < 0.05 ? (
                                      <div>
                                        Vì <strong>p-value ({pValue < 0.0001 ? "< 0.0001" : pValue.toFixed(4)}) &lt; 0.05</strong>, ta <strong>bác bỏ giả thuyết H0</strong> (giả thuyết hai chỉ tiêu độc lập), và chấp nhận giả thuyết đối H1.
                                        <br />
                                        <span className="inline-block mt-2 font-bold text-emerald-800">
                                          👉 Có sự tương quan liên hệ mật thiết, phụ thuộc và có ý nghĩa thống kê rất rõ rệt giữa hai đặc trưng "{col1}" và "{col2}".
                                        </span>
                                      </div>
                                    ) : (
                                      <div>
                                        Vì <strong>p-value ({pValue.toFixed(4)}) &ge; 0.05</strong>, <strong>chưa đủ bằng chứng bác bỏ giả thuyết H0</strong>.
                                        <br />
                                        <span className="inline-block mt-2 font-bold text-slate-600">
                                          👉 Hai đặc trưng "{col1}" và "{col2}" được coi là độc lập độc lập thống kê, không có mối quan hệ liên đới rõ rệt ở mức ý nghĩa 5%.
                                        </span>
                                      </div>
                                    )}
                                  </div>
                                </div>

                                <div className="bg-white/80 p-3 rounded-xl border border-slate-100 text-[10px] text-slate-500 space-y-1 leading-normal font-sans">
                                  <div><strong>Mẹo đọc nhanh:</strong> p-value dưới 0.05 biểu thị mối quan hệ vô cùng đáng tin cậy chứ không phải do ngẫu nhiên. Số càng nhỏ càng biểu thị tương quan chặt chẽ.</div>
                                </div>
                              </div>
                            </div>
                          </div>
                        ) : (
                          <div className="bg-slate-50 text-slate-500 rounded-xl p-4 text-xs italic text-center font-sans">
                            Không đủ số hàng hoặc cột (yêu cầu bảng chéo tối thiểu kích thước 2x2 và tổng số hộ &gt; 0) để chạy phép kiểm định độc lập Chi-Square.
                          </div>
                        )}
                      </div>

                      {/* INLINE MODAL SHOWING THE DETAILS OF CHOSEN CELL */}
                      {tqSelectedCell && (() => {
                        const { hang, cot } = tqSelectedCell;
                        
                        const matchingHouseholds = mainData.filter(row => {
                          let rVal = String(row[col1] ?? "").trim();
                          let cVal = String(row[col2] ?? "").trim();
                          if (!rVal) rVal = "(Trống)";
                          if (!cVal) cVal = "(Trống)";
                          
                          const matchHang = hang === null || rVal === hang;
                          const matchCot = cot === null || cVal === cot;
                          return matchHang && matchCot;
                        });

                        const filteredMatching = tqModalSearchTerm.trim()
                          ? matchingHouseholds.filter(row => {
                              const term = tqModalSearchTerm.toLowerCase();
                              return Object.values(row).some(val => String(val).toLowerCase().includes(term));
                            })
                          : matchingHouseholds;

                        const modalPageSize = 10;
                        const totalModalPages = Math.ceil(filteredMatching.length / modalPageSize) || 1;
                        const paginatedMatching = filteredMatching.slice(
                          (tqModalPage - 1) * modalPageSize,
                          (tqModalPage - 1) * modalPageSize + modalPageSize
                        );

                        const handleExportModalData = () => {
                          try {
                            const ws = XLSX.utils.json_to_sheet(filteredMatching);
                            const wb = XLSX.utils.book_new();
                            XLSX.utils.book_append_sheet(wb, ws, "Chi tiết ô tương quan");
                            const filePrefix = (hang ? `Hang_${hang}` : "All") + "_" + (cot ? `Cot_${cot}` : "All");
                            XLSX.writeFile(wb, `Danh_Sach_Ho_${filePrefix}_${new Date().toISOString().slice(0, 10)}.xlsx`);
                          } catch (e: any) {
                            alert("Lỗi xuất Excel: " + e.message);
                          }
                        };

                        const handleRedirectToMainTab = () => {
                          const matchingIndices: number[] = [];
                          mainData.forEach((row, idx) => {
                            let rVal = String(row[col1] ?? "").trim();
                            let cVal = String(row[col2] ?? "").trim();
                            if (!rVal) rVal = "(Trống)";
                            if (!cVal) cVal = "(Trống)";

                            const matchHang = hang === null || rVal === hang;
                            const matchCot = cot === null || cVal === cot;

                            if (matchHang && matchCot) {
                              matchingIndices.push(idx);
                            }
                          });

                          setRowIndicesFilter(matchingIndices);
                          
                          let label = "Tương quan: ";
                          if (hang && cot) {
                            label += `Hàng [${hang}] & Cột [${cot}]`;
                          } else if (hang) {
                            label += `Hàng [${hang}]`;
                          } else if (cot) {
                            label += `Cột [${cot}]`;
                          } else {
                            label += "Tổng cộng";
                          }
                          
                          setRowFilterLabel(label);
                          setViewPage(1);
                          setActiveTab("xemdulieu");
                          setTqSelectedCell(null);
                        };

                        return (
                          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade-in font-sans">
                            <div className="bg-white rounded-3xl max-w-4xl w-full shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[85vh]">
                              <div className="bg-slate-50 border-b border-slate-100 p-5 flex justify-between items-center shrink-0">
                                <div>
                                  <h4 className="text-sm font-black text-slate-800 uppercase tracking-wide flex items-center gap-2">
                                    <span>📋</span> DANH SÁCH CHI TIẾT ĐƠN VỊ HỘ KHẢO SÁT
                                  </h4>
                                  <p className="text-xs text-slate-500 mt-1">
                                    Phân loại: <span className="font-bold text-slate-700">{col1}</span> = {hang || "Tất cả"} × <span className="font-bold text-slate-700">{col2}</span> = {cot || "Tất cả"}
                                  </p>
                                </div>
                                <button 
                                  onClick={() => setTqSelectedCell(null)}
                                  className="text-slate-400 hover:text-slate-700 font-bold text-xl bg-transparent border-0 cursor-pointer p-1"
                                >
                                  ×
                                </button>
                              </div>

                              <div className="p-4 border-b border-slate-100 flex flex-col md:flex-row gap-3 items-center justify-between shrink-0 bg-slate-50/40">
                                <div className="relative w-full md:max-w-xs">
                                  <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
                                  <input
                                    type="text"
                                    value={tqModalSearchTerm}
                                    onChange={(e) => {
                                      setTqModalSearchTerm(e.target.value);
                                      setTqModalPage(1);
                                    }}
                                    placeholder="Tìm kiếm nhanh trong ô này..."
                                    className="w-full bg-white hover:border-slate-300 border border-slate-200 text-xs rounded-xl pl-9 pr-4 py-2 focus:outline-none focus:border-indigo-500 transition-all font-medium"
                                  />
                                </div>

                                <div className="flex gap-2 w-full md:w-auto">
                                  <button
                                    onClick={handleExportModalData}
                                    className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-4 py-2 rounded-xl transition-all shadow-sm border-0 flex items-center gap-1.5 cursor-pointer"
                                  >
                                    📥 Tải Excel ({filteredMatching.length.toLocaleString()} hộ)
                                  </button>
                                  <button
                                    onClick={handleRedirectToMainTab}
                                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold px-4 py-2 rounded-xl transition-all border border-slate-200 flex items-center gap-1.5 cursor-pointer"
                                  >
                                    📂 Xem nâng cao ở Tab Dữ Liệu
                                  </button>
                                </div>
                              </div>

                              <div className="flex-1 overflow-y-auto p-5 custom-scrollbar min-h-[250px]">
                                <div className="border border-slate-200 rounded-2xl overflow-hidden bg-white">
                                  <table className="w-full border-collapse text-left text-xs">
                                    <thead>
                                      <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider text-[10px]">
                                        <th className="p-3 w-12 text-center">STT</th>
                                        {columns.slice(0, 4).map(col => (
                                          <th key={col} className="p-3 font-sans truncate max-w-[150px]">{col}</th>
                                        ))}
                                        <th className="p-3 bg-indigo-50/50 text-indigo-700 font-bold max-w-[150px] truncate">{col1}</th>
                                        <th className="p-3 bg-emerald-50/50 text-emerald-700 font-bold max-w-[150px] truncate">{col2}</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {paginatedMatching.map((row, idx) => {
                                        const globalIdx = (tqModalPage - 1) * modalPageSize + idx + 1;
                                        return (
                                          <tr key={idx} className="border-b border-slate-100 hover:bg-slate-50/50 transition-colors">
                                            <td className="p-3 text-center font-mono text-slate-400">{globalIdx}</td>
                                            {columns.slice(0, 4).map(col => (
                                              <td key={col} className="p-3 truncate max-w-[150px] font-sans text-slate-700">
                                                {String(row[col] ?? "")}
                                              </td>
                                            ))}
                                            <td className="p-3 bg-indigo-50/20 font-bold text-slate-800">{String(row[col1] ?? "(Trống)")}</td>
                                            <td className="p-3 bg-emerald-50/20 font-bold text-slate-800">{String(row[col2] ?? "(Trống)")}</td>
                                          </tr>
                                        );
                                      })}

                                      {filteredMatching.length === 0 && (
                                        <tr>
                                          <td colSpan={columns.slice(0, 4).length + 3} className="p-10 text-center text-slate-400 italic font-sans">
                                            Không tìm thấy kết quả phù hợp trong danh mục này.
                                          </td>
                                        </tr>
                                      )}
                                    </tbody>
                                  </table>
                                </div>
                              </div>

                              <div className="bg-slate-50 border-t border-slate-100 p-4 flex flex-col sm:flex-row justify-between items-center gap-3 shrink-0">
                                <span className="text-[11px] text-slate-550 font-sans">
                                  Đang hiển thị <strong>{paginatedMatching.length}</strong> / <strong>{filteredMatching.length.toLocaleString()}</strong> hộ thỏa mãn
                                </span>

                                {totalModalPages > 1 && (
                                  <div className="flex items-center gap-1">
                                    <button
                                      disabled={tqModalPage === 1}
                                      onClick={() => setTqModalPage(prev => Math.max(1, prev - 1))}
                                      className="p-1.5 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                                    >
                                      ◀
                                    </button>
                                    <span className="text-[11px] font-bold text-slate-600 px-3 py-1 font-mono bg-white border border-slate-200 rounded-lg">
                                      Trang {tqModalPage} / {totalModalPages}
                                    </span>
                                    <button
                                      disabled={tqModalPage === totalModalPages}
                                      onClick={() => setTqModalPage(prev => Math.min(totalModalPages, prev + 1))}
                                      className="p-1.5 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                                    >
                                      ▶
                                    </button>
                                  </div>
                                )}

                                <button
                                  onClick={() => setTqSelectedCell(null)}
                                  className="bg-slate-800 hover:bg-slate-900 text-white text-xs font-bold px-5 py-2 rounded-xl transition-all border-0 cursor-pointer w-full sm:w-auto"
                                >
                                  Đóng
                                </button>
                              </div>
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  )}
                </div>
              );
            })()
          ) : (
            // ================= SUB TAB 2: TƯƠNG QUAN TUYẾN TÍNH =================
            (() => {
              const col1 = tqSelectedCol1 || (columns.length > 0 ? columns[0] : "");
              const col2 = tqSelectedCol2 || (columns.length > 1 ? columns[1] : (columns[0] || ""));

              if (!col1 || !col2) return <div className="text-xs text-slate-500 font-mono">Chưa chọn đủ hai cột dữ liệu.</div>;

              // Detect data types to determine if we should run pearson or grouped averages
              const pairs: { x: number; y: number; xRaw: any; yRaw: any }[] = [];
              let col1IsNumeric = true;
              let col2IsNumeric = true;

              const scanSample = mainData.slice(0, 100);
              let numericCount1 = 0;
              let numericCount2 = 0;

              scanSample.forEach(row => {
                const v1 = parseFloat(String(row[col1]).replace(/,/g, ""));
                const v2 = parseFloat(String(row[col2]).replace(/,/g, ""));
                if (!isNaN(v1)) numericCount1++;
                if (!isNaN(v2)) numericCount2++;
              });

              col1IsNumeric = numericCount1 > scanSample.length * 0.6;
              col2IsNumeric = numericCount2 > scanSample.length * 0.6;

              if (col1IsNumeric && col2IsNumeric) {
                mainData.forEach(row => {
                  const xVal = parseFloat(String(row[col1]).replace(/,/g, ""));
                  const yVal = parseFloat(String(row[col2]).replace(/,/g, ""));
                  if (!isNaN(xVal) && !isNaN(yVal)) {
                    pairs.push({ x: xVal, y: yVal, xRaw: row[col1], yRaw: row[col2] });
                  }
                });

                if (pairs.length < 3) {
                  return (
                    <div className="bg-amber-50 border border-amber-200 rounded-xl p-6 text-center text-xs text-amber-800 space-y-1">
                      <AlertTriangle className="w-6 h-6 text-amber-500 mx-auto" />
                      <p className="font-bold">Không đủ cặp giá trị số để thực hiện tương quan.</p>
                      <p className="text-[11px] text-slate-500">Cần ít nhất 3 dòng có giá trị số hợp lý ở cả hai cột {col1} và {col2}.</p>
                    </div>
                  );
                }

                const n = pairs.length;
                const sumX = pairs.reduce((sum, p) => sum + p.x, 0);
                const sumY = pairs.reduce((sum, p) => sum + p.y, 0);
                const sumXSq = pairs.reduce((sum, p) => sum + p.x * p.x, 0);
                const sumYSq = pairs.reduce((sum, p) => sum + p.y * p.y, 0);
                const sumXY = pairs.reduce((sum, p) => sum + p.x * p.y, 0);

                const num = n * sumXY - sumX * sumY;
                const den = Math.sqrt((n * sumXSq - sumX * sumX) * (n * sumYSq - sumY * sumY));

                const r = den !== 0 ? num / den : 0;
                
                let interp = "Không có tương quan";
                let interpColor = "text-slate-600 bg-slate-50 border-slate-250";
                const absR = Math.abs(r);
                if (absR >= 0.8) {
                  interp = r > 0 ? "Tương quan thuận rất mạnh (Rần đồng biến)" : "Tương quan nghịch rất mạnh (Rất nghịch biến)";
                  interpColor = "text-emerald-700 bg-emerald-50 border-emerald-200";
                } else if (absR >= 0.5) {
                  interp = r > 0 ? "Tương quan thuận khá mạnh (Đồng biến rõ rệt)" : "Tương quan nghịch khá mạnh (Nghịch biến rõ rệt)";
                  interpColor = "text-indigo-700 bg-indigo-50 border-indigo-200";
                } else if (absR >= 0.3) {
                  interp = r > 0 ? "Tương quan thuận vừa phải" : "Tương quan nghịch vừa phải";
                  interpColor = "text-sky-700 bg-sky-50 border-sky-200";
                } else if (absR > 0) {
                  interp = r > 0 ? "Tương quan thuận rất yếu" : "Tương quan nghịch rất yếu";
                  interpColor = "text-amber-700 bg-amber-50 border-amber-200";
                }

                const xVals = pairs.map(p => p.x);
                const yVals = pairs.map(p => p.y);
                const minX = Math.min(...xVals);
                const maxX = Math.max(...xVals);
                const minY = Math.min(...yVals);
                const maxY = Math.max(...yVals);

                const rangeX = maxX - minX || 1;
                const rangeY = maxY - minY || 1;

                return (
                  <div className="space-y-6">
                    <div className="flex flex-col md:flex-row gap-4">
                      <div className="space-y-1">
                        <label className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider font-mono">
                          Cột giá trị X (Số)
                        </label>
                        <select
                          value={tqSelectedCol1 || (columns.length > 0 ? columns[0] : "")}
                          onChange={(e) => setTqSelectedCol1(e.target.value)}
                          className="bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 cursor-pointer"
                        >
                          {columns.map((col) => (
                            <option key={col} value={col}>{col}</option>
                          ))}
                        </select>
                      </div>
                      <div className="space-y-1">
                        <label className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider font-mono">
                          Cột giá trị Y (Số)
                        </label>
                        <select
                          value={tqSelectedCol2 || (columns.length > 1 ? columns[1] : (columns[0] || ""))}
                          onChange={(e) => setTqSelectedCol2(e.target.value)}
                          className="bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 cursor-pointer"
                        >
                          {columns.map((col) => (
                            <option key={col} value={col}>{col}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div className={`border rounded-2xl p-5 flex flex-col md:flex-row md:items-center justify-between gap-4 transition-all ${interpColor}`}>
                      <div className="space-y-1">
                        <span className="text-[10px] font-bold uppercase tracking-wider font-mono opacity-80">HỆ SỐ TƯƠNG QUAN PEARSON (r)</span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-3xl font-black font-mono">{r.toFixed(4)}</span>
                          <span className="text-xs font-bold font-sans">({interp})</span>
                        </div>
                        <p className="text-[11px] font-medium opacity-90 font-sans">Tính toán dựa trên {n} cặp giá trị có định dạng số hợp lệ của hai cột.</p>
                      </div>
                      <div className="text-[10.5px] font-sans leading-normal max-w-sm border-l border-current/20 pl-4">
                        Hệ số r nằm trong khoảng [-1, 1]. Trị số gần 1 thể hiện đồng biến mạnh mẽ; trị số gần -1 thể hiện nghịch biến mạnh mẽ; trị số gần 0 thể hiện không có quan hệ tuyến tính rõ rệt.
                      </div>
                    </div>

                    <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-4 shadow-xl">
                      <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                        <h4 className="text-xs font-extrabold tracking-wider text-indigo-300 uppercase font-mono">
                          Biểu đồ phân tán đám mây điểm (Scatter Plot Diagram)
                        </h4>
                        <span className="text-[10px] font-mono text-slate-400 bg-slate-800 px-2 py-0.5 rounded">
                          X: {col1} | Y: {col2}
                        </span>
                      </div>

                      <div className="relative w-full h-80 bg-slate-950/80 rounded-xl border border-slate-800 overflow-hidden flex items-center justify-center">
                        <div className="absolute left-2.5 top-1/2 -rotate-90 origin-left text-[9px] font-mono text-slate-500 font-bold uppercase tracking-widest -translate-y-1/2">
                          {col2} →
                        </div>

                        <div className="absolute bottom-2 right-4 text-[9px] font-mono text-slate-500 font-bold uppercase tracking-widest">
                          {col1} →
                        </div>

                        <svg className="w-full h-full p-8" viewBox="0 0 500 250" preserveAspectRatio="none">
                          <line x1="0" y1="0" x2="500" y2="0" stroke="#1e293b" strokeWidth="1" strokeDasharray="3 3" />
                          <line x1="0" y1="62.5" x2="500" y2="62.5" stroke="#1e293b" strokeWidth="1" strokeDasharray="3 3" />
                          <line x1="0" y1="125" x2="500" y2="125" stroke="#1e293b" strokeWidth="1" strokeDasharray="3 3" />
                          <line x1="0" y1="187.5" x2="500" y2="187.5" stroke="#1e293b" strokeWidth="1" strokeDasharray="3 3" />
                          <line x1="0" y1="250" x2="500" y2="250" stroke="#1e293b" strokeWidth="1" />

                          <line x1="0" y1="0" x2="0" y2="250" stroke="#1e293b" strokeWidth="1" />
                          <line x1="125" y1="0" x2="125" y2="250" stroke="#1e293b" strokeWidth="1" strokeDasharray="3 3" />
                          <line x1="250" y1="0" x2="250" y2="250" stroke="#1e293b" strokeWidth="1" strokeDasharray="3 3" />
                          <line x1="375" y1="0" x2="375" y2="250" stroke="#1e293b" strokeWidth="1" strokeDasharray="3 3" />
                          <line x1="500" y1="0" x2="500" y2="250" stroke="#1e293b" strokeWidth="1" />

                          {(() => {
                            const avgX = sumX / n;
                            const avgY = sumY / n;
                            let numSlope = 0;
                            let denSlope = 0;
                            pairs.forEach(p => {
                              numSlope += (p.x - avgX) * (p.y - avgY);
                              denSlope += (p.x - avgX) * (p.x - avgX);
                            });
                            if (denSlope !== 0) {
                              const slope = numSlope / denSlope;
                              const intercept = avgY - slope * avgX;

                              const yStart = slope * minX + intercept;
                              const yEnd = slope * maxX + intercept;

                              const startXPix = 0;
                              const startYPix = 250 - ((yStart - minY) / rangeY) * 250;
                              const endXPix = 500;
                              const endYPix = 250 - ((yEnd - minY) / rangeY) * 250;

                              return (
                                <line 
                                  x1={startXPix} 
                                  y1={startYPix} 
                                  x2={endXPix} 
                                  y2={endYPix} 
                                  stroke="#6366f1" 
                                  strokeWidth="2.5" 
                                  title="Đường hồi quy tuyến tính" 
                                />
                              );
                            }
                            return null;
                          })()}

                          {pairs.slice(0, 300).map((p, idx) => {
                            const xPix = ((p.x - minX) / rangeX) * 500;
                            const yPix = 250 - ((p.y - minY) / rangeY) * 250;

                            return (
                              <circle
                                key={idx}
                                cx={xPix}
                                cy={yPix}
                                r="4.5"
                                className="fill-indigo-400 stroke-indigo-950 stroke-1 hover:fill-emerald-400 hover:r-6 cursor-pointer transition-all duration-100"
                              />
                            );
                          })}
                        </svg>

                        {pairs.length > 300 && (
                          <div className="absolute top-2.5 right-2.5 bg-slate-900/90 text-[8.5px] font-mono text-slate-400 border border-slate-800 rounded px-2 py-0.5">
                            Hiển thị 300 / {pairs.length} điểm biểu đồ tối ưu tốc độ
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                );
              } else {
                const categoryCol = col1IsNumeric ? col2 : col1;
                const numericCol = col1IsNumeric ? col1 : col2;

                const groupSums: Record<string, number> = {};
                const groupCounts: Record<string, number> = {};

                mainData.forEach(row => {
                  const catVal = String(row[categoryCol]).trim();
                  const numVal = parseFloat(String(row[numericCol]).replace(/,/g, ""));
                  if (catVal && !isNaN(numVal)) {
                    groupSums[catVal] = (groupSums[catVal] || 0) + numVal;
                    groupCounts[catVal] = (groupCounts[catVal] || 0) + 1;
                  }
                });

                const groupAverages = Object.entries(groupCounts)
                  .map(([category, count]) => {
                    const avg = groupSums[category] / count;
                    return { category, count, avg };
                  })
                  .sort((a, b) => b.avg - a.avg);

                return (
                  <div className="space-y-6">
                    <div className="flex flex-col md:flex-row gap-4">
                      <div className="space-y-1">
                        <label className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider font-mono">
                          Cột giá trị X (Phân loại/Số)
                        </label>
                        <select
                          value={tqSelectedCol1 || (columns.length > 0 ? columns[0] : "")}
                          onChange={(e) => setTqSelectedCol1(e.target.value)}
                          className="bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 cursor-pointer"
                        >
                          {columns.map((col) => (
                            <option key={col} value={col}>{col}</option>
                          ))}
                        </select>
                      </div>
                      <div className="space-y-1">
                        <label className="text-[11px] font-extrabold text-slate-500 uppercase tracking-wider font-mono">
                          Cột giá trị Y (Phân loại/Số)
                        </label>
                        <select
                          value={tqSelectedCol2 || (columns.length > 1 ? columns[1] : (columns[0] || ""))}
                          onChange={(e) => setTqSelectedCol2(e.target.value)}
                          className="bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 cursor-pointer"
                        >
                          {columns.map((col) => (
                            <option key={col} value={col}>{col}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div className="bg-slate-50 border border-slate-250 rounded-2xl p-5 space-y-2">
                      <h4 className="text-xs font-extrabold tracking-wider text-slate-700 uppercase font-mono">
                        Phân tích so sánh giá trị trung bình theo nhóm danh mục
                      </h4>
                      <p className="text-xs text-slate-500 leading-normal font-sans">
                        Hệ thống phát hiện thấy cột <strong className="text-slate-700">{categoryCol}</strong> là cột phân loại/văn bản và cột <strong className="text-slate-700">{numericCol}</strong> là cột chỉ số số. Sau đây là thống kê trung bình của <strong className="text-indigo-700">{numericCol}</strong> trên từng nhóm giá trị khác nhau:
                      </p>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 items-start">
                      <div className="bg-white border border-slate-200 rounded-2xl p-5 space-y-4">
                        <h4 className="text-xs font-extrabold tracking-wider text-slate-700 uppercase font-mono border-b border-slate-100 pb-2">
                          Biểu đồ trị trung bình của {numericCol} theo {categoryCol}
                        </h4>
                        {groupAverages.length === 0 ? (
                          <p className="text-xs text-slate-400 italic py-8 text-center">Không có dữ liệu so sánh.</p>
                        ) : (
                          <div className="space-y-4">
                            {groupAverages.slice(0, 15).map((item, idx) => {
                              const maxVal = groupAverages[0]?.avg || 1;
                              const pct = (item.avg / maxVal) * 100;
                              return (
                                <div key={idx} className="space-y-1 text-xs">
                                  <div className="flex justify-between text-[11px] font-medium text-slate-700">
                                    <span className="truncate max-w-[180px] font-sans font-bold" title={item.category}>
                                      {item.category}
                                    </span>
                                    <span className="font-mono text-emerald-700 font-bold">
                                      {item.avg.toLocaleString(undefined, { maximumFractionDigits: 2 })} (n={item.count})
                                    </span>
                                  </div>
                                  <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden border border-slate-200">
                                    <div 
                                      className="bg-gradient-to-r from-emerald-500 to-teal-500 h-2.5 rounded-full transition-all duration-500" 
                                      style={{ width: `${pct}%` }}
                                    ></div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>

                      <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                        <div className="p-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center">
                          <h4 className="text-xs font-extrabold tracking-wider text-slate-700 uppercase font-mono">
                            Bảng thống kê trị số trung bình
                          </h4>
                          <span className="text-[10px] bg-slate-200 text-slate-700 px-2 py-0.5 rounded font-bold font-mono">
                            {groupAverages.length} nhóm
                          </span>
                        </div>
                        <div className="max-h-[480px] overflow-y-auto custom-scrollbar">
                          <table className="w-full border-collapse text-[11px] text-left">
                            <thead>
                              <tr className="bg-slate-50 border-b border-slate-250 text-slate-500 font-mono font-bold sticky top-0">
                                <th className="p-3">Hạng</th>
                                <th className="p-3">Nhóm danh mục</th>
                                <th className="p-3 text-center">Số bản ghi (n)</th>
                                <th className="p-3 text-right">Trị trung bình ({numericCol})</th>
                              </tr>
                            </thead>
                            <tbody>
                              {groupAverages.map((item, idx) => (
                                <tr key={idx} className="border-b border-slate-100 hover:bg-slate-50 transition-colors">
                                  <td className="p-3 text-slate-400 font-mono font-bold">{idx + 1}</td>
                                  <td className="p-3 font-medium text-slate-800 truncate max-w-[150px]" title={item.category}>
                                    {item.category}
                                  </td>
                                  <td className="p-3 text-center font-mono font-medium text-slate-600">{item.count}</td>
                                  <td className="p-3 text-right font-mono font-bold text-emerald-700">
                                    {item.avg.toLocaleString(undefined, { maximumFractionDigits: 2 })}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              }
            })()
          )
        )}
      </div>
    </div>
  );
});
