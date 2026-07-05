import React, { useState } from "react";
import { GitMerge, FileUp, FolderOpen, Layers, Files, Trash2, HelpCircle, Info, Plus, CheckCircle, Settings, FileSpreadsheet } from "lucide-react";
import * as XLSX from "xlsx";
import { parseCSV } from "../utils/sharedHelpers";
import { MainDataInlinePreview } from "./MainDataInlinePreview";

interface FileMergerProps {
  mainData: any[];
  fileName: string;
  setMainData: (data: any[]) => void;
  setRawImportedData: (data: any[]) => void;
  setColumns: (cols: string[]) => void;
  setFileName: (name: string) => void;
  setCustomColConfigs: (configs: any[]) => void;
  setMapping: (mapping: any) => void;
  setLoading: (loading: boolean) => void;
  setProgress: (progress: number) => void;
  setStatusMessage: (msg: string) => void;
  onExportExcel: () => void;
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

export default function FileMerger({
  mainData,
  fileName,
  setMainData,
  setRawImportedData,
  setColumns,
  setFileName,
  setCustomColConfigs,
  setMapping,
  setLoading,
  setProgress,
  setStatusMessage,
  onExportExcel,
}: FileMergerProps) {
  // Chế độ ghép nối: 2 file (truyền thống) hoặc Hàng loạt thư mục (Yêu cầu mới)
  const [mergeMode, setMergeMode] = useState<"two-files" | "folder-batch">("two-files");

  // === CHẾ ĐỘ 1: GHÉP 2 FILE TRÁI & PHẢI ===
  const [leftData, setLeftData] = useState<any[]>([]);
  const [leftFileName, setLeftFileName] = useState<string>("");
  const [leftKey, setLeftKey] = useState<string>("");

  const [rightData, setRightData] = useState<any[]>([]);
  const [rightFileName, setRightFileName] = useState<string>("");
  const [rightKey, setRightKey] = useState<string>("");

  const [mergedResultData, setMergedResultData] = useState<any[] | null>(null);

  // === CHẾ ĐỘ 2: GHÉP THƯ MỤC / HÀNG LOẠT NHIỀU FILE ===
  const [batchFiles, setBatchFiles] = useState<{ name: string; size: number; data: any[]; cols: string[] }[]>([]);
  const [folderMergeKey, setFolderMergeKey] = useState<string>("");
  const [batchMergeMethod, setBatchMergeMethod] = useState<"join" | "append">("join");
  const [isFolderUpload, setIsFolderUpload] = useState<boolean>(false);

  // Đọc file đơn lẻ cho Chế độ 1 (Left & Right)
  const handleLocalFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: "left" | "right") => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setStatusMessage(`Đang đọc tệp tin: ${file.name}...`);

    const isCSV = file.name.toLowerCase().endsWith(".csv") || file.name.toLowerCase().endsWith(".txt");

    if (isCSV) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const text = evt.target?.result as string;
          if (!text) throw new Error("Không thể đọc nội dung tệp tin!");

          const data = parseCSV(text);
          if (data.length === 0) {
            alert("Tệp trống hoặc không chứa dữ liệu hợp lệ!");
            setLoading(false);
            return;
          }

          const keys = Object.keys(data[0] || {});
          if (type === "left") {
            setLeftData(data);
            setLeftFileName(file.name);
            if (keys.length > 0) setLeftKey(keys[0]);
          } else {
            setRightData(data);
            setRightFileName(file.name);
            if (keys.length > 0) setRightKey(keys[0]);
          }
          setStatusMessage(`Đã nạp thành công ${data.length} dòng.`);
        } catch (err: any) {
          alert("Lỗi khi đọc file CSV: " + err.message);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsText(file, "UTF-8");
    } else {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const arrayBuffer = evt.target?.result as ArrayBuffer;
          if (!arrayBuffer) throw new Error("Không thể đọc nội dung tệp tin!");

          const wb = XLSX.read(arrayBuffer, {
            type: "array",
            dense: true,
            cellFormula: false,
            cellHTML: false,
            cellStyles: false,
          });

          const wsName = wb.SheetNames[0];
          const ws = wb.Sheets[wsName];
          const data = XLSX.utils.sheet_to_json(ws) as any[];

          if (data.length === 0) {
            alert("Tệp Excel trống hoặc không chứa dữ liệu hợp lệ!");
            setLoading(false);
            return;
          }

          const keys = Object.keys(data[0] || {});
          if (type === "left") {
            setLeftData(data);
            setLeftFileName(file.name);
            if (keys.length > 0) setLeftKey(keys[0]);
          } else {
            setRightData(data);
            setRightFileName(file.name);
            if (keys.length > 0) setRightKey(keys[0]);
          }
          setStatusMessage(`Đã nạp thành công ${data.length} dòng.`);
        } catch (err: any) {
          alert("Lỗi khi đọc file Excel: " + err.message);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  // Nạp nhiều file hoặc cả Thư mục cho Chế độ 2
  const handleBatchFilesUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setLoading(true);
    setProgress(0);
    setStatusMessage(`Đang chuẩn bị nạp ${files.length} tệp tin...`);
    await sleep(250);

    const loadedFiles: typeof batchFiles = [];
    
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      // Bỏ qua các file rác hệ điều hành hoặc file tạm
      if (file.name.startsWith("._") || file.name.startsWith("~$") || file.name === ".DS_Store") continue;
      
      const ext = file.name.split(".").pop()?.toLowerCase();
      if (!["xlsx", "xls", "csv", "txt"].includes(ext || "")) continue;

      setStatusMessage(`Đang đọc tệp (${i + 1}/${files.length}): ${file.name}...`);
      setProgress(Math.floor((i / files.length) * 100));
      await sleep(10); // Cho phép UI render tiến trình
      
      const isCSV = ext === "csv" || ext === "txt";
      
      try {
        let fileData: any[] = [];
        if (isCSV) {
          const text = await new Promise<string>((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result as string);
            reader.onerror = reject;
            reader.readAsText(file, "UTF-8");
          });
          fileData = parseCSV(text);
        } else {
          const arrayBuffer = await new Promise<ArrayBuffer>((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result as ArrayBuffer);
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
          });
          const wb = XLSX.read(arrayBuffer, {
            type: "array",
            dense: true,
            cellFormula: false,
            cellHTML: false,
            cellStyles: false,
          });
          const wsName = wb.SheetNames[0];
          const ws = wb.Sheets[wsName];
          fileData = XLSX.utils.sheet_to_json(ws) as any[];
        }

        if (fileData.length > 0) {
          const cols = Object.keys(fileData[0] || {});
          loadedFiles.push({
            name: file.name,
            size: file.size,
            data: fileData,
            cols: cols
          });
        }
      } catch (err: any) {
        console.error(`Lỗi đọc file ${file.name}:`, err);
      }
    }

    if (loadedFiles.length === 0) {
      alert("Không tìm thấy tệp tin dữ liệu hợp lệ trong thư mục / danh sách chọn!");
      setLoading(false);
      return;
    }

    setBatchFiles(prev => [...prev, ...loadedFiles]);
    setLoading(false);
    setProgress(100);
    setStatusMessage(`Đã tải thành công ${loadedFiles.length} tệp dữ liệu vào danh sách.`);
    
    // Tự động tìm kiếm mã định danh phù hợp chung nhất
    const allColsList = loadedFiles.map(f => f.cols);
    if (allColsList.length > 0) {
      // Tìm các cột xuất hiện trong tất cả các file
      const commonCols = allColsList.reduce((acc, current) => acc.filter(x => current.includes(x)), allColsList[0]);
      if (commonCols.length > 0) {
        const bestGuess = commonCols.find(c => {
          const l = c.toLowerCase();
          return l.includes("id") || l.includes("mst") || l.includes("mã") || l.includes("key") || l.includes("code") || l.includes("định danh");
        });
        setFolderMergeKey(bestGuess || commonCols[0]);
      } else {
        // Dự phòng lấy cột đầu của file thứ nhất
        setFolderMergeKey(allColsList[0][0] || "");
      }
    }
  };

  // Ghép nối hai bảng Trái & Phải (Chế độ 1)
  const handleMerge = async () => {
    if (leftData.length === 0 || rightData.length === 0) {
      alert("Vui lòng tải đủ cả 2 bảng dữ liệu Trái & Phải!");
      return;
    }
    if (!leftKey || !rightKey) {
      alert("Vui lòng chọn cột khóa liên kết cho cả 2 bảng!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Bắt đầu ghép nối dữ liệu...");
    await sleep(200);

    const rightMap = new Map();
    const batchSize = Math.max(1, Math.floor(rightData.length / 10));

    for (let i = 0; i < rightData.length; i++) {
      const row = rightData[i];
      const kv = String(row[rightKey] || "").trim();
      if (kv) {
        rightMap.set(kv, row);
      }

      if (i % batchSize === 0 || i === rightData.length - 1) {
        const pct = Math.floor((i / rightData.length) * 40);
        setProgress(pct);
        setStatusMessage(`Đang lập chỉ mục bảng phải: ${i}/${rightData.length} dòng...`);
        await sleep(10);
      }
    }

    const mergedResults: any[] = [];
    const mainCols = Object.keys(leftData[0] || {});
    const rightCols = Object.keys(rightData[0] || {}).filter((c) => c !== rightKey);
    const stepSize = Math.max(1, Math.floor(leftData.length / 10));

    for (let j = 0; j < leftData.length; j++) {
      const leftRow = leftData[j];
      const matchKey = String(leftRow[leftKey] || "").trim();
      const matchedRight = rightMap.get(matchKey);

      const mergedRow = { ...leftRow };
      rightCols.forEach((rc) => {
        const finalColName = mainCols.includes(rc) ? `${rc}_Phai` : rc;
        mergedRow[finalColName] = matchedRight ? matchedRight[rc] : "";
      });

      mergedResults.push(mergedRow);

      if (j % stepSize === 0 || j === leftData.length - 1) {
        const pct = 40 + Math.floor((j / leftData.length) * 60);
        setProgress(pct);
        setStatusMessage(`Đang ánh xạ & ghép dòng: ${j}/${leftData.length} dòng...`);
        await sleep(10);
      }
    }

    const mergedCols = Object.keys(mergedResults[0] || {});
    setMainData(mergedResults);
    setRawImportedData(mergedResults);
    setColumns(mergedCols);
    setFileName(`GhepNoi_${leftFileName.replace(/\.[^/.]+$/, "")}_vs_${rightFileName.replace(/\.[^/.]+$/, "")}.xlsx`);
    setMergedResultData(mergedResults);

    setMapping({
      mota: "",
      manganh: "",
      xa: "",
      doanhthu: "",
      laodong: "",
      idCol: "",
    });

    const initMergedConfigs = mergedCols.map((c) => {
      return {
        originalName: c,
        use: true,
        newName: c,
        role: "" as any,
      };
    });
    setCustomColConfigs(initMergedConfigs);

    setProgress(100);
    setStatusMessage(`Ghép nối thành công hoàn tất! Thu được ${mergedResults.length} dòng dữ liệu.`);
    await sleep(400);
    setLoading(false);
  };

  // Ghép nối hàng loạt / thư mục nhiều file (Chế độ 2)
  const handleBatchMerge = async () => {
    if (batchFiles.length < 2) {
      alert("Vui lòng tải lên hoặc nạp ít nhất 2 tệp tin trong danh sách!");
      return;
    }

    if (batchMergeMethod === "join" && !folderMergeKey) {
      alert("Vui lòng chỉ định Mã Định Danh chung để so khớp các cột dữ liệu!");
      return;
    }

    setLoading(true);
    setProgress(0);
    setStatusMessage("Bắt đầu thực hiện ghép nối thư mục / hàng loạt...");
    await sleep(300);

    let finalMerged: any[] = [];

    if (batchMergeMethod === "append") {
      // 1. GHÉP NỐI TIẾP (APPEND ROWS / UNION)
      setStatusMessage("Đang ghép nối dòng hàng loạt (Union)...");
      let totalRowsParsed = 0;
      
      batchFiles.forEach((file, idx) => {
        // Chuẩn hóa tên cột của từng dòng
        file.data.forEach(row => {
          finalMerged.push({
            ...row,
            Nguon_File: file.name
          });
        });
        totalRowsParsed += file.data.length;
        setProgress(Math.floor(((idx + 1) / batchFiles.length) * 100));
      });
    } else {
      // 2. GHÉP SONG SONG THEO MÃ ĐỊNH DANH (JOIN COLUMNS ON KEY)
      setStatusMessage(`Đang gộp cột song song dựa theo khóa "${folderMergeKey}"...`);
      
      const masterKeyMap = new Map<string, any>();
      
      batchFiles.forEach((file, fileIdx) => {
        const fileShortName = file.name.replace(/\.[^/.]+$/, "");
        
        file.data.forEach((row) => {
          const rawKeyVal = row[folderMergeKey];
          if (rawKeyVal === undefined || rawKeyVal === null) return;
          const keyString = String(rawKeyVal).trim();
          if (!keyString) return;

          let masterRow = masterKeyMap.get(keyString);
          if (!masterRow) {
            masterRow = { [folderMergeKey]: keyString };
            masterKeyMap.set(keyString, masterRow);
          }

          // Gộp các thuộc tính từ tệp hiện tại vào dòng gốc
          Object.keys(row).forEach((col) => {
            if (col === folderMergeKey) return;
            
            // Nếu cột đã tồn tại từ một file trước đó, đổi tên để tránh ghi đè dữ liệu quý báu
            let finalColName = col;
            if (masterRow[col] !== undefined) {
              finalColName = `${col}_${fileShortName}`;
            }
            masterRow[finalColName] = row[col];
          });
        });
        
        setProgress(Math.floor(((fileIdx + 1) / batchFiles.length) * 100));
      });

      finalMerged = Array.from(masterKeyMap.values());
    }

    if (finalMerged.length === 0) {
      alert("Lỗi: Không thu được dòng dữ liệu nào sau khi gộp!");
      setLoading(false);
      return;
    }

    const mergedCols = Object.keys(finalMerged[0] || {});
    setMainData(finalMerged);
    setRawImportedData(finalMerged);
    setColumns(mergedCols);
    setFileName(`GhepNoiThurMuc_HangLoat_${batchFiles.length}_files.xlsx`);
    setMergedResultData(finalMerged);

    setMapping({
      mota: "",
      manganh: "",
      xa: "",
      doanhthu: "",
      laodong: "",
      idCol: "",
    });

    const initMergedConfigs = mergedCols.map((c) => {
      return {
        originalName: c,
        use: true,
        newName: c,
        role: "" as any,
      };
    });
    setCustomColConfigs(initMergedConfigs);

    setProgress(100);
    setStatusMessage(`Ghép nối thành công ${batchFiles.length} file trong thư mục! Tổng thu được ${finalMerged.length} dòng.`);
    await sleep(400);
    setLoading(false);
  };

  const removeBatchFile = (index: number) => {
    setBatchFiles(prev => prev.filter((_, i) => i !== index));
  };

  const clearAllBatchFiles = () => {
    setBatchFiles([]);
    setFolderMergeKey("");
  };

  return (
    <div className="space-y-6 animate-fade-in font-sans">
      
      {/* Pills menu switcher between Two Files vs Folder Batch */}
      <div className="flex border-b border-slate-200">
        <button
          onClick={() => setMergeMode("two-files")}
          className={`py-3 px-6 text-sm font-bold border-b-2 transition-all flex items-center gap-2 cursor-pointer ${
            mergeMode === "two-files"
              ? "border-indigo-600 text-indigo-600 bg-indigo-50/40"
              : "border-transparent text-slate-500 hover:text-slate-700 hover:bg-slate-50"
          }`}
        >
          <GitMerge className="w-4 h-4" />
          Ghép 2 Biểu Mẫu (Left Join)
        </button>
        <button
          onClick={() => setMergeMode("folder-batch")}
          className={`py-3 px-6 text-sm font-bold border-b-2 transition-all flex items-center gap-2 cursor-pointer ${
            mergeMode === "folder-batch"
              ? "border-indigo-600 text-indigo-600 bg-indigo-50/40"
              : "border-transparent text-slate-500 hover:text-slate-700 hover:bg-slate-50"
          }`}
        >
          <FolderOpen className="w-4 h-4 text-emerald-500" />
          Ghép Thư Mục / Nhiều Tệp Tự Động
          <span className="bg-emerald-100 text-emerald-800 text-[10px] px-1.5 py-0.5 rounded-full font-mono font-bold">Mới</span>
        </button>
      </div>

      {/* CHẾ ĐỘ 1: GHÉP 2 FILE TRÁI - PHẢI */}
      {mergeMode === "two-files" && (
        <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-6 text-slate-800">
          <div>
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <GitMerge className="w-5 h-5 text-indigo-500" /> GHÉP NỐI HAI BIỂU DỮ LIỆU (LEFT JOIN)
            </h3>
            <p className="text-xs text-slate-500 mt-1">
              Ánh xạ, gộp và bổ sung thêm các cột chỉ tiêu từ bảng bên Phải vào bảng bên Trái dựa trên mã định danh chung (như mã số thuế, mã xã, mã huyện...).
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* BẢNG TRÁI */}
            <div className="bg-slate-50 border border-slate-100 p-5 rounded-xl space-y-4">
              <span className="text-xs font-bold text-indigo-600 tracking-wider uppercase font-mono block">
                📁 1. BẢNG DỮ LIỆU TRÁI (LÀM GỐC)
              </span>

              <label className="bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 shadow-sm font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center gap-2 cursor-pointer w-full justify-center">
                <FileUp className="w-4 h-4 text-indigo-500" /> CHỌN BẢNG TRÁI (.xlsx, .xls, .csv)
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) => handleLocalFileUpload(e, "left")}
                  className="hidden"
                />
              </label>

              {mainData && mainData.length > 0 && (
                <button
                  onClick={() => {
                    setLeftData(mainData);
                    setLeftFileName(fileName || "Dữ liệu chính hệ thống");
                    const keys = Object.keys(mainData[0] || {});
                    if (keys.length > 0) {
                      setLeftKey(keys[0]);
                    }
                  }}
                  type="button"
                  className="w-full bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-[11px] px-3 py-2 rounded-xl border border-indigo-200 transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-sm"
                >
                  📥 SỬ DỤNG FILE CHÍNH ĐÃ NẠP
                </button>
              )}

              {leftFileName && (
                <div className="bg-white border border-slate-200 p-3 rounded-lg text-xs flex justify-between items-center text-slate-700 shadow-sm">
                  <span className="truncate max-w-[200px]" title={leftFileName}>📄 {leftFileName}</span>
                  <span className="font-mono text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded">{leftData.length} dòng</span>
                </div>
              )}

              {leftData.length > 0 && (
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-700 block">Chọn Cột Khóa Bảng Trái:</label>
                  <select
                    value={leftKey}
                    onChange={(e) => setLeftKey(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800"
                  >
                    <option value="">-- Chọn cột khóa liên kết --</option>
                    {Object.keys(leftData[0] || {}).map((c) => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>
              )}
            </div>

            {/* BẢNG PHẢI */}
            <div className="bg-slate-50 border border-slate-100 p-5 rounded-xl space-y-4">
              <span className="text-xs font-bold text-emerald-600 tracking-wider uppercase font-mono block">
                📁 2. BẢNG DỮ LIỆU PHẢI (ÁNH XẠ)
              </span>

              <label className="bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 shadow-sm font-bold text-xs px-4 py-2.5 rounded-xl transition-all flex items-center gap-2 cursor-pointer w-full justify-center">
                <FileUp className="w-4 h-4 text-emerald-500" /> CHỌN BẢNG PHẢI (.xlsx, .xls, .csv)
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) => handleLocalFileUpload(e, "right")}
                  className="hidden"
                />
              </label>

              {rightFileName && (
                <div className="bg-white border border-slate-200 p-3 rounded-lg text-xs flex justify-between items-center text-slate-700 shadow-sm">
                  <span className="truncate max-w-[200px]" title={rightFileName}>📄 {rightFileName}</span>
                  <span className="font-mono text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded">{rightData.length} dòng</span>
                </div>
              )}

              {rightData.length > 0 && (
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-700 block">Chọn Cột Khóa Bảng Phải:</label>
                  <select
                    value={rightKey}
                    onChange={(e) => setRightKey(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800"
                  >
                    <option value="">-- Chọn cột khóa liên kết --</option>
                    {Object.keys(rightData[0] || {}).map((c) => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>
              )}
            </div>
          </div>

          <button
            onClick={handleMerge}
            className="w-full bg-gradient-to-r from-indigo-500 to-indigo-600 hover:from-indigo-600 hover:to-indigo-700 text-white font-bold text-xs px-6 py-3.5 rounded-xl border-b-4 border-indigo-700 active:scale-95 transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer"
          >
            <GitMerge className="w-4 h-4" /> THỰC HIỆN GHÉP NỐI & NẠP VÀO HỆ THỐNG XEM CHÍNH
          </button>
        </div>
      )}

      {/* CHẾ ĐỘ 2: GHÉP THƯ MỤC / HÀNG LOẠT NHIỀU FILE TỰ ĐỘNG */}
      {mergeMode === "folder-batch" && (
        <div className="bg-white border border-slate-200 shadow-sm rounded-2xl p-6 space-y-6 text-slate-800 animate-fade-in">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div>
              <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <FolderOpen className="w-5 h-5 text-emerald-500" /> TỰ ĐỘNG GHÉP THƯ MỤC & NHIỀU FILE HÀNG LOẠT
              </h3>
              <p className="text-xs text-slate-500 mt-1">
                Đặt các tệp khảo sát, báo cáo tháng/quý vào chung thư mục hoặc chọn nhiều file cùng lúc. Hệ thống tự động phân tích và hỏi gộp theo mã định danh.
              </p>
            </div>
            {batchFiles.length > 0 && (
              <button
                onClick={clearAllBatchFiles}
                className="text-xs font-semibold py-1 px-3 border border-red-200 hover:border-red-300 bg-red-50 text-red-700 rounded-xl transition cursor-pointer"
              >
                🧹 Xóa danh sách tệp ({batchFiles.length})
              </button>
            )}
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {/* Cột 1: Drag & drop / Upload Folder & Config */}
            <div className="space-y-4 md:col-span-1">
              <span className="text-xs font-extrabold text-emerald-700 tracking-wider uppercase font-mono block">
                📁 1. NẠP DỮ LIỆU TỪ THƯ MỤC
              </span>

              {/* Toggle upload mode: Folder vs Multiple Files */}
              <div className="flex bg-slate-100 p-1 rounded-xl gap-1">
                <button
                  type="button"
                  onClick={() => setIsFolderUpload(false)}
                  className={`flex-1 py-1 px-2 text-[11px] font-bold rounded-lg transition ${
                    !isFolderUpload ? "bg-white text-slate-800 shadow-xs" : "text-slate-500 hover:text-slate-700"
                  }`}
                >
                  Chọn nhiều Files
                </button>
                <button
                  type="button"
                  onClick={() => setIsFolderUpload(true)}
                  className={`flex-1 py-1 px-2 text-[11px] font-bold rounded-lg transition ${
                    isFolderUpload ? "bg-white text-slate-800 shadow-xs" : "text-slate-500 hover:text-slate-700"
                  }`}
                >
                  Tải cả Thư mục (Folder)
                </button>
              </div>

              <div className="relative group">
                <input
                  type="file"
                  multiple
                  {...(isFolderUpload ? { webkitdirectory: "", directory: "" } as any : {})}
                  onChange={handleBatchFilesUpload}
                  className="hidden"
                  id="batch-files-upload-input"
                />
                <label
                  htmlFor="batch-files-upload-input"
                  className="flex flex-col items-center justify-center border-2 border-dashed border-slate-300 hover:border-emerald-500 rounded-2xl p-8 text-center cursor-pointer transition bg-slate-50/50 hover:bg-emerald-50/25 shadow-xs"
                >
                  <FolderOpen className="w-10 h-10 text-slate-400 group-hover:text-emerald-500 mb-2 animate-pulse" />
                  <span className="text-xs font-bold text-slate-700 group-hover:text-emerald-600">
                    {isFolderUpload ? "Nhấp chọn Thư mục" : "Bấm chọn nhiều file cùng lúc"}
                  </span>
                  <span className="text-[10px] text-slate-400 mt-1">Hỗ trợ Excel (.xlsx, .xls) & CSV</span>
                </label>
              </div>

              {/* Phương thức gộp */}
              {batchFiles.length > 0 && (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3 shadow-xs">
                  <div className="space-y-1">
                    <label className="text-[11px] font-extrabold text-slate-600 uppercase block">Cách thức ghép nối:</label>
                    <div className="space-y-1.5">
                      <label className="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer">
                        <input
                          type="radio"
                          name="batchMergeMethod"
                          checked={batchMergeMethod === "join"}
                          onChange={() => setBatchMergeMethod("join")}
                          className="text-indigo-600 focus:ring-indigo-500 h-3.5 w-3.5"
                        />
                        Ghép song song (Gộp cột theo Mã)
                      </label>
                      <label className="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer">
                        <input
                          type="radio"
                          name="batchMergeMethod"
                          checked={batchMergeMethod === "append"}
                          onChange={() => setBatchMergeMethod("append")}
                          className="text-indigo-600 focus:ring-indigo-500 h-3.5 w-3.5"
                        />
                        Ghép nối tiếp (Gộp dòng - Append)
                      </label>
                    </div>
                  </div>

                  {batchMergeMethod === "join" && (
                    <div className="space-y-1 pt-2 border-t border-slate-200">
                      <label className="text-[11px] font-extrabold text-slate-600 uppercase block">🔑 Cột Mã định danh định danh chung:</label>
                      <select
                        value={folderMergeKey}
                        onChange={(e) => setFolderMergeKey(e.target.value)}
                        className="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-800 focus:ring-1 focus:ring-emerald-500"
                      >
                        <option value="">-- Chọn mã định danh chung --</option>
                        {Array.from(new Set(batchFiles.flatMap(f => f.cols))).map(col => (
                          <option key={col} value={col}>{col}</option>
                        ))}
                      </select>
                      <p className="text-[10px] text-slate-400">
                        Ví dụ: MST, Số định danh, Mã địa bàn... dùng để đối khớp hàng ngang.
                      </p>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Cột 2 & 3: Danh sách tệp tin đã nạp */}
            <div className="md:col-span-2 space-y-3">
              <span className="text-xs font-extrabold text-slate-500 tracking-wider uppercase font-mono block">
                📋 Danh sách tệp tin đang chờ ghép nối ({batchFiles.length})
              </span>

              {batchFiles.length === 0 ? (
                <div className="border border-dashed border-slate-200 rounded-2xl p-12 text-center text-slate-400 flex flex-col items-center justify-center bg-slate-50/20">
                  <Files className="w-10 h-10 text-slate-300 mb-2" />
                  <p className="text-xs font-medium">Chưa có tệp tin nào được nạp.</p>
                  <p className="text-[10px] text-slate-400 mt-1">Vui lòng chọn thư mục chứa các tệp tin cùng loại ở cột bên trái.</p>
                </div>
              ) : (
                <div className="border border-slate-200 rounded-xl bg-slate-50/30 overflow-hidden shadow-xs max-h-[320px] overflow-y-auto">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100 border-b border-slate-200 text-slate-600 font-bold select-none">
                        <th className="p-3">Tên file</th>
                        <th className="p-3 w-24 text-right">Kích cỡ</th>
                        <th className="p-3 w-28 text-right">Số dòng nạp</th>
                        <th className="p-3 w-16 text-center">Xóa</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 bg-white">
                      {batchFiles.map((f, idx) => (
                        <tr key={f.name + "_" + idx} className="hover:bg-slate-50/50 transition">
                          <td className="p-3 font-semibold text-slate-700 truncate max-w-[240px]" title={f.name}>
                            📄 {f.name}
                          </td>
                          <td className="p-3 text-right text-slate-500 font-mono">
                            {(f.size / 1024).toFixed(1)} KB
                          </td>
                          <td className="p-3 text-right font-bold text-slate-600 font-mono">
                            {f.data.length.toLocaleString()} dòng
                          </td>
                          <td className="p-3 text-center">
                            <button
                              onClick={() => removeBatchFile(idx)}
                              type="button"
                              className="text-red-500 hover:text-red-700 p-1 hover:bg-red-50 rounded transition cursor-pointer"
                              title="Loại bỏ file này"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {batchFiles.length >= 2 && (
                <div className="pt-2 animate-fade-in">
                  <button
                    onClick={handleBatchMerge}
                    className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-6 py-4 rounded-xl border-b-4 border-emerald-800 active:scale-95 transition-all shadow-md flex items-center justify-center gap-2 cursor-pointer"
                  >
                    <CheckCircle className="w-4 h-4 animate-bounce" />
                    BẮT ĐẦU TỰ ĐỘNG GHÉP {batchFiles.length} TỆP TRONG THƯ MỤC
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* KẾT QUẢ HIỂN THỊ */}
      {mergedResultData && mergedResultData.length > 0 && (
        <MainDataInlinePreview
          data={mergedResultData}
          columns={Object.keys(mergedResultData[0] || {})}
          title={mergeMode === "two-files" ? "KẾT QUẢ GHÉP NỐI DỮ LIỆU (LEFT JOIN)" : "KẾT QUẢ GHÉP NỐI THƯ MỤC / HÀNG LOẠT"}
          subtitle={`Hệ thống đã hợp nhất thành công! Tổng số thu được: ${mergedResultData.length} dòng dữ liệu.`}
          onExportExcel={onExportExcel}
        />
      )}

      {/* HIỂN THỊ BẢNG TRÁI ĐỂ QUAN SÁT */}
      {mergeMode === "two-files" && leftData.length > 0 && (
        <MainDataInlinePreview
          data={leftData}
          columns={Object.keys(leftData[0] || {})}
          title="DỮ LIỆU NGUỒN BẢNG TRÁI"
          subtitle="Xem trước bảng trái đang được chọn làm cơ sở dữ liệu gốc."
        />
      )}
    </div>
  );
}
