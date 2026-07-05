import React, { useState, useMemo, useEffect } from "react";
import { Sliders, AlertTriangle, Search, Download, Check, Settings, Info, RefreshCw, FileSpreadsheet, Play, Upload, Palette } from "lucide-react";

interface FrequencyAnalysisProps {
  mainData: any[];
  columns: string[];
}

export const FrequencyAnalysis = React.memo(function FrequencyAnalysis({
  mainData,
  columns
}: FrequencyAnalysisProps) {
  // Trạng thái phân tích tần suất đơn giản (cũ)
  const [tsSelectedCol, setTsSelectedCol] = useState<string>("");
  const [isTsCalculated, setIsTsCalculated] = useState<boolean>(false);

  // Trạng thái cho "BẢNG PHÂN PHỐI TẦN SUẤT ĐỊA BÀN XÃ (CÓ / KHÔNG)" - MỚI THEO YÊU CẦU
  const [configProvCol, setConfigProvCol] = useState("");
  const [configDistCol, setConfigDistCol] = useState("");
  const [configCommCol, setConfigCommCol] = useState("");
  const [configCommCodeCol, setConfigCommCodeCol] = useState("");
  const [configSheetCol, setConfigSheetCol] = useState("");
  const [configIndicatorCol, setConfigIndicatorCol] = useState("");

  const [selectedProv, setSelectedProv] = useState("Tất cả");
  const [selectedDist, setSelectedDist] = useState("Tất cả");
  const [selectedSheet, setSelectedSheet] = useState("Tất cả");

  const [showConfigPanel, setShowConfigPanel] = useState(false);
  const [communeSearchTerm, setCommuneSearchTerm] = useState("");
  const [communeRows, setCommuneRows] = useState<any[]>([]);
  const [hasRunAnalysis, setHasRunAnalysis] = useState(false);

  // Trạng thái cho theme màu sắc và tải file trong THỐNG KÊ TẦN SUẤT ĐỊA BÀN XÃ
  const [showThemePanel, setShowThemePanel] = useState(false);
  const [customTheme, setCustomTheme] = useState({
    headerBg: "#f1f5f9", // Slate-100 default
    headerText: "#334155", // Slate-700
    totalBg: "#fef3c7", // Amber-50
    totalText: "#78350f", // Amber-900
    rowEvenBg: "#ffffff",
    rowOddBg: "#f8fafc",
    borderColor: "#e2e8f0"
  });

  const colorPresets = [
    {
      name: "Mặc định Slate",
      headerBg: "#f1f5f9",
      headerText: "#334155",
      totalBg: "#fef3c7",
      totalText: "#78350f",
      rowEvenBg: "#ffffff",
      rowOddBg: "#f8fafc",
    },
    {
      name: "Xanh lá Excel",
      headerBg: "#107c41",
      headerText: "#ffffff",
      totalBg: "#e8f5e9",
      totalText: "#1b5e20",
      rowEvenBg: "#ffffff",
      rowOddBg: "#f1f8f3",
    },
    {
      name: "Xanh đại dương",
      headerBg: "#1e3a8a",
      headerText: "#ffffff",
      totalBg: "#dbeafe",
      totalText: "#1e40af",
      rowEvenBg: "#ffffff",
      rowOddBg: "#f0f5ff",
    },
    {
      name: "Đỏ nổi bật",
      headerBg: "#991b1b",
      headerText: "#ffffff",
      totalBg: "#fee2e2",
      totalText: "#991b1b",
      rowEvenBg: "#ffffff",
      rowOddBg: "#fff5f5",
    },
    {
      name: "Đá phiến tối giản",
      headerBg: "#334155",
      headerText: "#ffffff",
      totalBg: "#f1f5f9",
      totalText: "#0f172a",
      rowEvenBg: "#ffffff",
      rowOddBg: "#fafafa",
    }
  ];

  const isLightColor = (hex: string) => {
    const color = hex.replace("#", "");
    if (color.length !== 6) return true;
    const r = parseInt(color.substring(0, 2), 16);
    const g = parseInt(color.substring(2, 4), 16);
    const b = parseInt(color.substring(4, 6), 16);
    const brightness = (r * 299 + g * 587 + b * 114) / 1000;
    return brightness > 155;
  };

  // Default selected column if not set for basic analysis
  const col = useMemo(() => {
    return tsSelectedCol || (columns.length > 0 ? columns[0] : "");
  }, [tsSelectedCol, columns]);

  // Compute basic frequency (cũ)
  const frequencyResult = useMemo(() => {
    if (!mainData || mainData.length === 0 || !col) {
      return { sortedFreqs: [], uniqueCount: 0, blankCount: 0, topValue: null, totalCount: 0 };
    }

    const freqs: Record<string, number> = {};
    let totalCount = 0;
    let blankCount = 0;

    mainData.forEach(row => {
      const val = row[col];
      if (val === undefined || val === null || String(val).trim() === "") {
        blankCount++;
      } else {
        const valStr = String(val).trim();
        freqs[valStr] = (freqs[valStr] || 0) + 1;
        totalCount++;
      }
    });

    const sortedFreqs = Object.entries(freqs)
      .map(([value, count]) => ({ value, count, pct: (count / totalCount) * 100 }))
      .sort((a, b) => b.count - a.count);

    const topValue = sortedFreqs[0] || null;
    const uniqueCount = sortedFreqs.length;

    return {
      sortedFreqs,
      uniqueCount,
      blankCount,
      topValue,
      totalCount
    };
  }, [mainData, col]);

  const { sortedFreqs, uniqueCount, blankCount, topValue } = frequencyResult;

  // Tự động nhận diện các cột dựa trên tên cột
  const detectColumns = () => {
    let provCol = "";
    let distCol = "";
    let commCol = "";
    let commCodeCol = "";
    let sheetCol = "";

    columns.forEach(c => {
      const low = c.toLowerCase();
      if (!provCol && (low.includes("tỉnh") || low.includes("tinh") || low.includes("thành phố") || low.includes("thanh pho") || low === "tp" || low.includes("province"))) {
        provCol = c;
      }
      if (!distCol && (low.includes("huyện") || low.includes("huyen") || low.includes("quận") || low.includes("quan") || low.includes("district"))) {
        distCol = c;
      }
      if (!commCol && (low.includes("xã") || low.includes("xa") || low.includes("phường") || low.includes("phuong") || low.includes("commune")) && !low.includes("mã") && !low.includes("ma")) {
        commCol = c;
      }
      if (!commCodeCol && (low.includes("mã xã") || low.includes("ma xa") || low.includes("maxa") || low.includes("mã_xã") || low.includes("ma_xa") || low.includes("commune_code") || low.includes("commune code"))) {
        commCodeCol = c;
      }
      if (!sheetCol && (low.includes("loại phiếu") || low.includes("loai phieu") || low.includes("phiếu") || low.includes("phieu") || low.includes("form") || low.includes("sheet"))) {
        sheetCol = c;
      }
    });

    return {
      provCol: provCol || "",
      distCol: distCol || "",
      commCol: commCol || (columns.find(c => c.toLowerCase().includes("xã") || c.toLowerCase().includes("xa")) || ""),
      commCodeCol: commCodeCol || "",
      sheetCol: sheetCol || ""
    };
  };

  // Khởi tạo cột cấu hình tự động khi columns thay đổi không đè cấu hình thủ công của người dùng
  useEffect(() => {
    if (columns.length > 0) {
      const detected = detectColumns();
      setConfigProvCol(prev => (prev && columns.includes(prev)) ? prev : detected.provCol);
      setConfigDistCol(prev => (prev && columns.includes(prev)) ? prev : detected.distCol);
      setConfigCommCol(prev => (prev && columns.includes(prev)) ? prev : detected.commCol);
      setConfigCommCodeCol(prev => (prev && columns.includes(prev)) ? prev : detected.commCodeCol);
      setConfigSheetCol(prev => (prev && columns.includes(prev)) ? prev : detected.sheetCol);
      
      setConfigIndicatorCol(prev => {
        if (prev && columns.includes(prev)) return prev;
        const currentProv = configProvCol || detected.provCol;
        const currentDist = configDistCol || detected.distCol;
        const currentComm = configCommCol || detected.commCol;
        const currentCommCode = configCommCodeCol || detected.commCodeCol;
        const currentSheet = configSheetCol || detected.sheetCol;
        
        const filterCol = columns.find(c => 
          c !== currentProv && 
          c !== currentDist && 
          c !== currentComm && 
          c !== currentCommCode && 
          c !== currentSheet
        );
        return filterCol || columns[0];
      });
    }
  }, [columns]);

  // Tạo danh sách dropdown dựa trên dữ liệu thực tế
  const provincesList = useMemo(() => {
    if (!configProvCol || !mainData || mainData.length === 0) return [];
    const set = new Set<string>();
    mainData.forEach(r => {
      const val = r[configProvCol];
      if (val !== undefined && val !== null && String(val).trim() !== "") {
        set.add(String(val).trim());
      }
    });
    return Array.from(set).sort();
  }, [mainData, configProvCol]);

  const districtsList = useMemo(() => {
    if (!configDistCol || !mainData || mainData.length === 0) return [];
    const set = new Set<string>();
    mainData.forEach(r => {
      if (configProvCol && selectedProv !== "Tất cả" && String(r[configProvCol]).trim() !== selectedProv) {
        return;
      }
      const val = r[configDistCol];
      if (val !== undefined && val !== null && String(val).trim() !== "") {
        set.add(String(val).trim());
      }
    });
    return Array.from(set).sort();
  }, [mainData, configDistCol, configProvCol, selectedProv]);

  const sheetsList = useMemo(() => {
    if (!configSheetCol || !mainData || mainData.length === 0) return [];
    const set = new Set<string>();
    mainData.forEach(r => {
      const val = r[configSheetCol];
      if (val !== undefined && val !== null && String(val).trim() !== "") {
        set.add(String(val).trim());
      }
    });
    return Array.from(set).sort();
  }, [mainData, configSheetCol]);

  // Khởi chạy tính toán Thống kê Tần suất theo địa bàn Xã
  const handleRunCommuneAnalysis = () => {
    if (!mainData || mainData.length === 0) {
      alert("Không có dữ liệu khảo sát trong bộ nhớ.");
      return;
    }
    if (!configCommCol) {
      alert("Vui lòng thiết lập cột Địa bàn Xã ở mục Cấu hình!");
      setShowConfigPanel(true);
      return;
    }
    if (!configIndicatorCol) {
      alert("Vui lòng chọn cột Chỉ tiêu cần thống kê!");
      return;
    }

    // 1. Lọc theo Tỉnh, Huyện, Loại phiếu
    let filtered = [...mainData];
    if (configProvCol && selectedProv !== "Tất cả") {
      filtered = filtered.filter(r => String(r[configProvCol]).trim() === selectedProv);
    }
    if (configDistCol && selectedDist !== "Tất cả") {
      filtered = filtered.filter(r => String(r[configDistCol]).trim() === selectedDist);
    }
    if (configSheetCol && selectedSheet !== "Tất cả") {
      filtered = filtered.filter(r => String(r[configSheetCol]).trim() === selectedSheet);
    }

    // 2. Gom nhóm theo Xã
    const groups: Record<string, { maXa: string; tenXa: string; rows: any[] }> = {};

    filtered.forEach(row => {
      const rawXa = String(row[configCommCol] || "").trim();
      if (!rawXa) return;

      const rawMaXa = configCommCodeCol ? String(row[configCommCodeCol] || "").trim() : "";
      const key = `${rawMaXa}_${rawXa}`;
      if (!groups[key]) {
        groups[key] = {
          maXa: rawMaXa || "00000",
          tenXa: rawXa,
          rows: []
        };
      }
      groups[key].rows.push(row);
    });

    // Định nghĩa tiêu chí khẳng định "Có"
    const affirmativeWords = ["có", "co", "1", "yes", "y", "đúng", "dung", "đạt", "dat", "true", "đạt chuẩn"];

    // 3. Tính toán tần suất Có/Không cho từng Xã
    const computedRows: any[] = Object.values(groups).map((g) => {
      const tongSo = g.rows.length;
      let co = 0;
      g.rows.forEach(r => {
        const valStr = String(r[configIndicatorCol] || "").trim().toLowerCase();
        const isAffirmative = affirmativeWords.some(w => {
          if (w === "1") return valStr === "1";
          if (w === "y") return valStr === "y";
          return valStr === w || valStr.includes(w);
        });
        if (isAffirmative) {
          co++;
        }
      });

      const khong = tongSo - co;

      return {
        maXa: g.maXa === "00000" ? String(Math.floor(10000 + Math.random() * 90000)) : g.maXa, // Tạo mã xã ngẫu nhiên nếu trống
        tenXa: g.tenXa,
        tongSo,
        co,
        khong,
        isTotalRow: false
      };
    });

    // Sắp xếp các xã theo bảng chữ cái tiếng Việt
    computedRows.sort((a, b) => a.tenXa.localeCompare(b.tenXa, "vi"));

    // 4. Tính toán dòng Tổng cộng Huyện ở đầu
    const grandTotal = computedRows.reduce((acc, curr) => {
      acc.tongSo += curr.tongSo;
      acc.co += curr.co;
      acc.khong += curr.khong;
      return acc;
    }, { tongSo: 0, co: 0, khong: 0 });

    let totalLabel = "Huyện Hưng Hà";
    if (selectedDist !== "Tất cả") {
      totalLabel = selectedDist;
      if (!totalLabel.toLowerCase().includes("huyện") && !totalLabel.toLowerCase().includes("quận") && !totalLabel.toLowerCase().includes("thị xã")) {
        totalLabel = "Huyện " + totalLabel;
      }
    } else if (districtsList.length > 0) {
      totalLabel = districtsList[0];
      if (!totalLabel.toLowerCase().includes("huyện") && !totalLabel.toLowerCase().includes("quận") && !totalLabel.toLowerCase().includes("thị xã")) {
        totalLabel = "Huyện " + totalLabel;
      }
    }

    const finalRows = [
      {
        maXa: "00000",
        tenXa: totalLabel,
        tongSo: grandTotal.tongSo,
        co: grandTotal.co,
        khong: grandTotal.khong,
        isTotalRow: true
      },
      ...computedRows
    ];

    setCommuneRows(finalRows);
    setHasRunAnalysis(true);
  };

  // Cho phép chỉnh sửa trực tiếp con số trên bảng và tự động cộng dồn ngược
  const handleCellChangeCommune = (idx: number, field: "tongSo" | "co" | "khong", valStr: string) => {
    const cleanNumStr = valStr.replace(/[^0-9]/g, "");
    const numericVal = cleanNumStr === "" ? 0 : parseInt(cleanNumStr) || 0;

    setCommuneRows(prev => {
      const updated = [...prev];
      const row = { ...updated[idx] };

      row[field] = numericVal;

      // Ràng buộc tính toán Tổng số = Có + Không
      if (field === "co") {
        row.khong = Math.max(0, row.tongSo - numericVal);
      } else if (field === "khong") {
        row.co = Math.max(0, row.tongSo - numericVal);
      } else if (field === "tongSo") {
        row.khong = Math.max(0, numericVal - row.co);
      }

      updated[idx] = row;

      // Tính toán lại dòng Tổng cộng đầu tiên (Mã xã "00000")
      const totalRow = { ...updated[0] };
      let sumTong = 0;
      let sumCo = 0;
      let sumKhong = 0;

      for (let i = 1; i < updated.length; i++) {
        sumTong += updated[i].tongSo || 0;
        sumCo += updated[i].co || 0;
        sumKhong += updated[i].khong || 0;
      }

      totalRow.tongSo = sumTong;
      totalRow.co = sumCo;
      totalRow.khong = sumKhong;
      updated[0] = totalRow;

      return updated;
    });
  };

  // Tải bảng tần suất địa bàn Xã xuống dưới dạng file Excel
  const handleExportCommuneExcel = () => {
    if (communeRows.length === 0) {
      alert("Không có số liệu tần suất để xuất!");
      return;
    }

    try {
      import("xlsx").then((XLSX) => {
        const excelData = communeRows.map((r, i) => ({
          "STT": r.isTotalRow ? "" : i,
          "Mã xã": r.maXa,
          "Tên": r.tenXa,
          "Tổng số": r.tongSo,
          "Có": r.co,
          "Không": r.khong
        }));

        const ws = XLSX.utils.json_to_sheet(excelData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Tần suất Xã");
        
        // Căn lề và phong cách cơ bản
        XLSX.writeFile(wb, `Thong_Ke_Tan_Suat_Xa_${configIndicatorCol.replace(/[^a-zA-Z0-9_]/g, "_")}_Da_Sua.xlsx`);
      });
    } catch (err: any) {
      alert("Lỗi xuất Excel: " + err.message);
    }
  };

  // Đọc file Excel tần suất xã có sẵn, tự động ánh xạ cột và màu sắc của file
  const handleCommuneExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        import("xlsx").then((XLSX) => {
          const workbook = XLSX.read(bstr, { type: "binary", cellStyles: true });
          const wsname = workbook.SheetNames[0];
          const ws = workbook.Sheets[wsname];
          
          const rawData = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1 });
          if (rawData.length < 2) {
            alert("File Excel trống hoặc không đúng định dạng!");
            return;
          }

          // Phân tích hàng tiêu đề để tìm cột
          let headerIdx = -1;
          let cols: string[] = [];
          for (let r = 0; r < Math.min(rawData.length, 10); r++) {
            const row = rawData[r];
            if (row && Array.isArray(row) && row.some((cell: any) => {
              const str = String(cell || "").toLowerCase();
              return str.includes("mã xã") || str.includes("tên") || str.includes("tổng") || str.includes("có") || str.includes("không");
            })) {
              headerIdx = r;
              cols = row.map((cell: any) => String(cell || "").trim());
              break;
            }
          }

          if (headerIdx === -1) {
            headerIdx = 0;
            cols = Array.isArray(rawData[0]) ? rawData[0].map((cell: any) => String(cell || "").trim()) : [];
          }

          let maXaColIdx = -1;
          let tenXaColIdx = -1;
          let tongSoColIdx = -1;
          let coColIdx = -1;
          let khongColIdx = -1;

          cols.forEach((colName, idx) => {
            const lower = colName.toLowerCase();
            if (lower.includes("mã xã") || lower.includes("ma xa") || lower.includes("mã") || lower.includes("code") || lower.includes("mã địa bàn")) {
              maXaColIdx = idx;
            } else if (lower.includes("tên") || lower.includes("ten") || lower.includes("xã") || lower.includes("commune") || lower.includes("địa bàn")) {
              tenXaColIdx = idx;
            } else if (lower.includes("tổng") || lower.includes("tong") || lower.includes("total") || lower.includes("số lượng")) {
              tongSoColIdx = idx;
            } else if (lower.includes("có") || lower.includes("co") || lower.includes("yes") || lower.includes("đạt")) {
              coColIdx = idx;
            } else if (lower.includes("không") || lower.includes("khong") || lower.includes("no") || lower.includes("chưa đạt")) {
              khongColIdx = idx;
            }
          });

          // Cài đặt dự phòng chỉ số cột nếu không nhận diện được
          if (maXaColIdx === -1) maXaColIdx = 1;
          if (tenXaColIdx === -1) tenXaColIdx = 2;
          if (tongSoColIdx === -1) tongSoColIdx = 3;
          if (coColIdx === -1) coColIdx = 4;
          if (khongColIdx === -1) khongColIdx = 5;

          const rows: any[] = [];
          for (let r = headerIdx + 1; r < rawData.length; r++) {
            const rowData = rawData[r];
            if (!rowData || !Array.isArray(rowData) || rowData.length === 0) continue;

            const nameVal = tenXaColIdx < rowData.length ? String(rowData[tenXaColIdx] || "").trim() : "";
            if (!nameVal) continue;

            const codeVal = maXaColIdx < rowData.length ? String(rowData[maXaColIdx] || "").trim() : "";
            const totalVal = tongSoColIdx < rowData.length ? parseInt(String(rowData[tongSoColIdx]).replace(/[^0-9]/g, "")) || 0 : 0;
            const coVal = coColIdx < rowData.length ? parseInt(String(rowData[coColIdx]).replace(/[^0-9]/g, "")) || 0 : 0;
            const khongVal = khongColIdx < rowData.length ? parseInt(String(rowData[khongColIdx]).replace(/[^0-9]/g, "")) || 0 : 0;

            const isTotalRow = nameVal.toLowerCase().includes("tổng") || nameVal.toLowerCase().includes("cộng") || codeVal === "00000";

            rows.push({
              maXa: codeVal || (isTotalRow ? "00000" : String(Math.floor(10000 + Math.random() * 90000))),
              tenXa: nameVal,
              tongSo: totalVal,
              co: coVal,
              khong: khongVal,
              isTotalRow: isTotalRow
            });
          }

          if (rows.length === 0) {
            alert("Không tìm thấy dữ liệu dòng nào phù hợp trong file Excel!");
            return;
          }

          const hasTotalRow = rows.some(r => r.isTotalRow);
          let finalRows = [...rows];
          if (!hasTotalRow) {
            const sumTong = rows.reduce((acc, c) => acc + c.tongSo, 0);
            const sumCo = rows.reduce((acc, c) => acc + c.co, 0);
            const sumKhong = rows.reduce((acc, c) => acc + c.khong, 0);
            const generatedTotal = {
              maXa: "00000",
              tenXa: "TỔNG CỘNG",
              tongSo: sumTong,
              co: sumCo,
              khong: sumKhong,
              isTotalRow: true
            };
            finalRows = [generatedTotal, ...rows];
          } else {
            const totalRowIndex = finalRows.findIndex(r => r.isTotalRow);
            if (totalRowIndex > 0) {
              const [tRow] = finalRows.splice(totalRowIndex, 1);
              finalRows.unshift(tRow);
            }
          }

          // Cố gắng quét màu nền từ workbook (Yêu cầu "áp dụng luôn theo màu của file tài liệu")
          let detectedHeaderBg = "";
          let detectedTotalBg = "";

          const refRange = ws['!ref'];
          if (refRange) {
            const decoded = XLSX.utils.decode_range(refRange);
            // 1. Quét màu của dòng tiêu đề
            for (let c = decoded.s.c; c <= decoded.e.c; c++) {
              const cellRef = XLSX.utils.encode_cell({ r: headerIdx, c: c });
              const cell = ws[cellRef];
              if (cell && cell.s && cell.s.fill && cell.s.fill.fgColor) {
                const colorObj = cell.s.fill.fgColor;
                if (colorObj.rgb) {
                  let rgbStr = String(colorObj.rgb);
                  if (rgbStr.length === 8) rgbStr = rgbStr.substring(2);
                  detectedHeaderBg = "#" + rgbStr;
                  break;
                }
              }
            }

            // 2. Quét màu của dòng tổng cộng
            const totalRowOffset = hasTotalRow ? rows.findIndex(r => r.isTotalRow) : -1;
            if (totalRowOffset !== -1) {
              const totalRowIdxInExcel = totalRowOffset + headerIdx + 1;
              for (let c = decoded.s.c; c <= decoded.e.c; c++) {
                const cellRef = XLSX.utils.encode_cell({ r: totalRowIdxInExcel, c: c });
                const cell = ws[cellRef];
                if (cell && cell.s && cell.s.fill && cell.s.fill.fgColor) {
                  const colorObj = cell.s.fill.fgColor;
                  if (colorObj.rgb) {
                    let rgbStr = String(colorObj.rgb);
                    if (rgbStr.length === 8) rgbStr = rgbStr.substring(2);
                    detectedTotalBg = "#" + rgbStr;
                    break;
                  }
                }
              }
            }
          }

          // Cập nhật theme màu nếu phát hiện được
          setCustomTheme(prev => ({
            ...prev,
            headerBg: detectedHeaderBg || prev.headerBg,
            headerText: detectedHeaderBg ? (isLightColor(detectedHeaderBg) ? "#1e293b" : "#ffffff") : prev.headerText,
            totalBg: detectedTotalBg || prev.totalBg,
            totalText: detectedTotalBg ? (isLightColor(detectedTotalBg) ? "#7c2d12" : "#ffffff") : prev.totalText,
          }));

          setCommuneRows(finalRows);
          setHasRunAnalysis(true);
          alert(`Đã nạp thành công ${finalRows.length - 1} xã và tự động áp dụng cấu trúc màu sắc từ file Excel tài liệu!`);
        });
      } catch (err: any) {
        alert("Lỗi đọc file Excel: " + err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Bộ lọc tìm kiếm xã trên bảng
  const filteredCommuneRows = useMemo(() => {
    if (!communeSearchTerm.trim()) return communeRows;
    const term = communeSearchTerm.toLowerCase();
    
    // Luôn giữ dòng tổng cộng ở đầu, lọc các dòng xã còn lại
    const totalRow = communeRows.find(r => r.isTotalRow);
    const otherRows = communeRows.filter(r => !r.isTotalRow && (
      String(r.tenXa).toLowerCase().includes(term) || 
      String(r.maXa).toLowerCase().includes(term)
    ));

    return totalRow ? [totalRow, ...otherRows] : otherRows;
  }, [communeRows, communeSearchTerm]);

  if (mainData.length === 0) {
    return (
      <div className="bg-slate-50 rounded-2xl p-12 text-center text-xs text-amber-700 border border-slate-200 font-condensed space-y-4 shadow-sm font-sans">
        <div className="flex justify-center">
          <AlertTriangle className="w-8 h-8 text-amber-500 animate-bounce" />
        </div>
        <div>
          <h3 className="text-sm font-bold uppercase tracking-wider text-slate-800">CHƯA CÓ DỮ LIỆU NGUỒN KHẢO SÁT</h3>
          <p className="text-xs text-slate-500 max-w-md mx-auto mt-1 font-sans">
            Hãy nạp file Excel hoặc CSV danh sách đơn vị cơ sở khảo sát (doanh nghiệp/hộ cá thể) của bạn tại Trang Chủ trước để hệ thống tiến hành tính toán phân tích và thiết lập mẫu.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-8 animate-fade-in font-sans text-slate-800">
      
      {/* ================== PHÂN HỆ 1: PHÂN TÍCH TẦN SUẤT ĐƠN GIẢN (HIỆN TẠI) ================== */}
      <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
          <div className="space-y-1">
            <h3 className="text-sm font-extrabold uppercase tracking-wider text-slate-800 flex items-center gap-2">
              <Sliders className="w-4 h-4 text-indigo-600 animate-pulse" />
              PHÂN TÍCH TẦN SUẤT XUẤT HIỆN DỮ LIỆU
            </h3>
            <p className="text-xs text-slate-500">
              Thống kê mức độ tập trung, tỷ lệ phần trăm và biểu đồ phân bổ của một cột dữ liệu bất kỳ.
            </p>
          </div>

          <div className="flex items-center gap-2">
            <label className="text-xs font-bold text-slate-600 font-mono">CHỌN CỘT PHÂN TÍCH:</label>
            <select
              value={col}
              onChange={(e) => {
                setTsSelectedCol(e.target.value);
                setIsTsCalculated(false);
              }}
              className="bg-slate-50 hover:bg-slate-100 border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 cursor-pointer"
            >
              {columns.map((columnName) => (
                <option key={columnName} value={columnName}>{columnName}</option>
              ))}
            </select>
          </div>
        </div>

        {!isTsCalculated ? (
          <div className="bg-slate-900 text-slate-100 rounded-2xl p-8 border border-slate-800 flex flex-col items-center justify-center text-center space-y-4 max-w-xl mx-auto my-6 shadow-md animate-fade-in">
            <div className="bg-indigo-500/10 text-indigo-400 p-4 rounded-full border border-indigo-500/20">
              <Sliders className="w-8 h-8 text-indigo-400 animate-pulse" />
            </div>
            <div className="space-y-1.5">
              <h4 className="text-sm font-black tracking-wider uppercase text-slate-200">
                SẴN SÀNG PHÂN TÍCH TẦN SUẤT
              </h4>
              <p className="text-xs text-slate-400 max-w-sm">
                Cột đang chọn: <strong className="text-indigo-400 font-mono">[{col || "Chưa chọn"}]</strong>. Nhấn nút bên dưới để khởi chạy quét và lập thống kê phân bổ.
              </p>
            </div>
            <button
              onClick={() => setIsTsCalculated(true)}
              className="bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs py-2.5 px-6 rounded-xl border-0 cursor-pointer shadow transition-all active:scale-95 flex items-center gap-1.5"
            >
              ⚡ KHỞI CHẠY TÍNH TOÁN TẦN SUẤT
            </button>
          </div>
        ) : (
          <div className="space-y-6 animate-fade-in">
            {/* Summary widgets */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-1">
                <span className="text-[10px] text-slate-400 font-bold tracking-wider font-mono">TỔNG SỐ DÒNG</span>
                <p className="text-lg font-black text-slate-800 font-mono">{mainData.length} dòng</p>
              </div>
              <div className="bg-indigo-50/50 border border-indigo-100 rounded-2xl p-4 space-y-1">
                <span className="text-[10px] text-indigo-500 font-bold tracking-wider font-mono">SỐ GIÁ TRỊ KHÁC BIỆT</span>
                <p className="text-lg font-black text-indigo-700 font-mono">{uniqueCount} giá trị</p>
              </div>
              <div className="bg-emerald-50/50 border border-emerald-100 rounded-2xl p-4 space-y-1">
                <span className="text-[10px] text-emerald-600 font-bold tracking-wider font-mono">GIÁ TRỊ XUẤT HIỆN NHIỀU NHẤT</span>
                <p className="text-xs font-bold text-emerald-800 truncate" title={topValue?.value || "N/A"}>
                  {topValue ? topValue.value : "N/A"}
                </p>
              </div>
              <div className="bg-amber-50/50 border border-amber-100 rounded-2xl p-4 space-y-1">
                <span className="text-[10px] text-amber-600 font-bold tracking-wider font-mono">TẦN SUẤT TRỐNG (BLANK)</span>
                <p className="text-lg font-black text-amber-700 font-mono">
                  {blankCount} ô ({((blankCount / mainData.length) * 100).toFixed(1)}%)
                </p>
              </div>
            </div>

            {/* Interactive Graph + Table splits */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
              {/* Left: Beautiful CSS/SVG Bar Chart */}
              <div className="lg:col-span-7 bg-slate-50 rounded-2xl border border-slate-200 p-5 space-y-4">
                <h4 className="text-xs font-extrabold tracking-wider text-slate-700 uppercase font-mono border-b border-slate-200 pb-2">
                  Biểu đồ tần suất 15 nhóm dẫn đầu (%)
                </h4>

                {sortedFreqs.length === 0 ? (
                  <p className="text-xs text-slate-400 italic py-8 text-center">Không có dữ liệu phân bổ.</p>
                ) : (
                  <div className="space-y-4">
                    {sortedFreqs.slice(0, 15).map((item, idx) => {
                      const pctVal = item.pct;
                      return (
                        <div key={idx} className="space-y-1 text-xs">
                          <div className="flex justify-between text-[11px] font-medium text-slate-700">
                            <span className="truncate max-w-[200px] font-sans font-bold" title={item.value}>
                              {item.value}
                            </span>
                            <span className="font-mono text-slate-500 font-bold">
                              {item.count} lần ({pctVal.toFixed(2)}%)
                            </span>
                          </div>
                          <div className="w-full bg-slate-250 rounded-full h-2.5 overflow-hidden shadow-xs border border-slate-200">
                            <div 
                              className="bg-gradient-to-r from-indigo-500 via-indigo-600 to-sky-500 h-2.5 rounded-full transition-all duration-500" 
                              style={{ width: `${pctVal}%` }}
                            ></div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Right: Data Table of counts */}
              <div className="lg:col-span-5 bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm">
                <div className="p-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center">
                  <h4 className="text-xs font-extrabold tracking-wider text-slate-700 uppercase font-mono">
                    Danh sách phân phối đầy đủ
                  </h4>
                  <span className="text-[10px] bg-slate-200 text-slate-700 px-2 py-0.5 rounded font-bold font-mono">
                    {sortedFreqs.length} dòng
                  </span>
                </div>
                <div className="max-h-[480px] overflow-y-auto custom-scrollbar">
                  <table className="w-full border-collapse text-[11px] text-left">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-250 text-slate-500 font-mono font-bold sticky top-0">
                        <th className="p-3">Hạng</th>
                        <th className="p-3">Giá trị</th>
                        <th className="p-3 text-center">Số lần</th>
                        <th className="p-3 text-right">Tỷ lệ %</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedFreqs.map((item, idx) => (
                        <tr key={idx} className="border-b border-slate-100 hover:bg-slate-50 transition-colors">
                          <td className="p-3 text-slate-400 font-mono font-bold">{idx + 1}</td>
                          <td className="p-3 font-medium text-slate-800 truncate max-w-[150px]" title={item.value}>
                            {item.value}
                          </td>
                          <td className="p-3 text-center font-mono font-bold text-indigo-700">{item.count}</td>
                          <td className="p-3 text-right font-mono text-slate-550">{item.pct.toFixed(2)}%</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>


      {/* ================== PHÂN HỆ 2: BẢNG PHÂN PHỐI TẦN SUẤT ĐỊA BÀN XÃ (CÓ / KHÔNG) - MỚI THEO ẢNH CHỤP ================== */}
      <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xl space-y-6" id="commune_frequency_analysis_section">
        
        {/* TITLE BLOCK */}
        <div className="border-b border-slate-200 pb-4 flex flex-col md:flex-row justify-between md:items-center gap-4">
          <div className="space-y-1">
            <h3 className="text-sm font-black uppercase tracking-widest text-indigo-900 flex items-center gap-2">
              📊 THỐNG KÊ TẦN SUẤT ĐỊA BÀN XÃ (CÓ / KHÔNG)
            </h3>
            <p className="text-xs text-slate-500">
              Phân nhóm dữ liệu theo từng xã, đếm số lượng câu trả lời khẳng định (Có, Đúng, Đạt) đối với chỉ tiêu được chọn.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {/* Import File & Theme button */}
            <button
              onClick={() => {
                setShowThemePanel(!showThemePanel);
                setShowConfigPanel(false);
              }}
              className={`text-xs font-bold py-1.5 px-3.5 rounded-xl border transition-all flex items-center gap-1.5 cursor-pointer shadow-sm ${
                showThemePanel 
                  ? "bg-emerald-55 border-emerald-200 text-emerald-700" 
                  : "bg-white border-slate-300 hover:bg-slate-50 text-slate-700"
              }`}
            >
              <Upload className="w-3.5 h-3.5" />
              📂 Nạp file & Màu sắc
            </button>

            {/* Toggle config columns */}
            <button
              onClick={() => {
                setShowConfigPanel(!showConfigPanel);
                setShowThemePanel(false);
              }}
              className={`text-xs font-bold py-1.5 px-3.5 rounded-xl border transition-all flex items-center gap-1.5 cursor-pointer shadow-sm ${
                showConfigPanel 
                  ? "bg-indigo-50 border-indigo-200 text-indigo-700" 
                  : "bg-white border-slate-300 hover:bg-slate-50 text-slate-700"
              }`}
            >
              <Settings className="w-3.5 h-3.5" />
              {showConfigPanel ? "Đóng cài đặt cột" : "⚙️ Cấu hình cột địa bàn"}
            </button>
          </div>
        </div>

        {/* COLLAPSIBLE UPLOAD & COLORS PANEL */}
        {showThemePanel && (
          <div className="bg-slate-50 rounded-xl p-5 border border-slate-200 space-y-6 animate-fade-in font-sans">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              
              {/* LEFT: EXCEL UPLOAD FIELD */}
              <div className="space-y-3 border-r border-slate-200 pr-0 md:pr-6">
                <div className="flex items-center gap-1.5 text-xs font-extrabold text-emerald-800 uppercase tracking-wider">
                  <Upload className="w-4 h-4 text-emerald-600" />
                  NẠP SỐ LIỆU TẦN SUẤT XÃ TỪ FILE EXCEL CÓ SẴN
                </div>
                <p className="text-[11px] text-slate-500 leading-relaxed">
                  Bạn có thể tải trực tiếp file Excel tần suất xã có sẵn (thay vì bấm "Xem KQ" để tổng hợp tự động từ danh sách khảo sát thô). File tải lên chỉ cần chứa các cột dạng: <strong>STT, Mã xã, Tên xã, Tổng số, Có, Không</strong>.
                </p>
                
                <div className="relative group">
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    onChange={handleCommuneExcelUpload}
                    className="hidden"
                    id="commune-excel-upload-input"
                  />
                  <label
                    htmlFor="commune-excel-upload-input"
                    className="flex flex-col items-center justify-center border-2 border-dashed border-slate-300 hover:border-emerald-500 rounded-xl p-6 text-center cursor-pointer transition bg-white hover:bg-emerald-50/20 shadow-xs"
                  >
                    <Upload className="w-8 h-8 text-slate-400 group-hover:text-emerald-500 animate-pulse mb-2" />
                    <span className="text-xs font-bold text-slate-700 group-hover:text-emerald-600">Bấm để chọn file Excel tần suất xã</span>
                    <span className="text-[10px] text-slate-400 mt-1">Đồng bộ tự động dữ liệu và màu sắc từ file</span>
                  </label>
                </div>
              </div>

              {/* RIGHT: THEME & COLOR PICKER */}
              <div className="space-y-4">
                <div className="flex items-center gap-1.5 text-xs font-extrabold text-indigo-800 uppercase tracking-wider">
                  <Palette className="w-4 h-4 text-indigo-600" />
                  TÙY CHỈNH MÀU SẮC BẢNG THEO TÀI LIỆU
                </div>
                <p className="text-[11px] text-slate-500">
                  Chọn tông màu nhanh từ các preset có sẵn hoặc tự định nghĩa mã màu Hex để bảng hiển thị đồng bộ tuyệt đối với tệp tài liệu báo cáo của bạn.
                </p>

                {/* Color presets list */}
                <div className="flex flex-wrap gap-2 pb-2">
                  {colorPresets.map((preset) => (
                    <button
                      key={preset.name}
                      onClick={() => {
                        setCustomTheme({
                          ...customTheme,
                          headerBg: preset.headerBg,
                          headerText: preset.headerText,
                          totalBg: preset.totalBg,
                          totalText: preset.totalText,
                          rowEvenBg: preset.rowEvenBg,
                          rowOddBg: preset.rowOddBg,
                        });
                      }}
                      className="text-[11px] font-semibold py-1 px-2.5 rounded-lg border border-slate-205 bg-white hover:bg-slate-50 transition shadow-xs flex items-center gap-1.5 cursor-pointer"
                    >
                      <span 
                        className="w-3 h-3 rounded-full border border-slate-300"
                        style={{ backgroundColor: preset.headerBg }}
                      ></span>
                      {preset.name}
                    </button>
                  ))}
                </div>

                {/* Advanced Color pickers */}
                <div className="grid grid-cols-2 gap-3 bg-white p-3 rounded-xl border border-slate-250">
                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider block font-mono">Màu nền Header:</label>
                    <div className="flex items-center gap-1.5">
                      <input 
                        type="color" 
                        value={customTheme.headerBg} 
                        onChange={(e) => {
                          const val = e.target.value;
                          setCustomTheme(prev => ({
                            ...prev,
                            headerBg: val,
                            headerText: isLightColor(val) ? "#1e293b" : "#ffffff"
                          }));
                        }}
                        className="w-6 h-6 rounded cursor-pointer border border-slate-300 p-0"
                      />
                      <input 
                        type="text" 
                        value={customTheme.headerBg}
                        onChange={(e) => {
                          const val = e.target.value;
                          if (val.startsWith("#") && val.length <= 7) {
                            setCustomTheme(prev => ({
                              ...prev,
                              headerBg: val,
                              headerText: isLightColor(val) ? "#1e293b" : "#ffffff"
                            }));
                          }
                        }}
                        className="bg-slate-50 border border-slate-250 rounded px-2 py-1 text-[11px] font-mono text-slate-700 w-20"
                      />
                    </div>
                  </div>

                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider block font-mono">Chữ Header:</label>
                    <div className="flex items-center gap-1.5">
                      <input 
                        type="color" 
                        value={customTheme.headerText} 
                        onChange={(e) => setCustomTheme(prev => ({ ...prev, headerText: e.target.value }))}
                        className="w-6 h-6 rounded cursor-pointer border border-slate-300 p-0"
                      />
                      <input 
                        type="text" 
                        value={customTheme.headerText}
                        onChange={(e) => {
                          const val = e.target.value;
                          if (val.startsWith("#") && val.length <= 7) {
                            setCustomTheme(prev => ({ ...prev, headerText: val }));
                          }
                        }}
                        className="bg-slate-50 border border-slate-250 rounded px-2 py-1 text-[11px] font-mono text-slate-700 w-20"
                      />
                    </div>
                  </div>

                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider block font-mono">Nền Tổng cộng:</label>
                    <div className="flex items-center gap-1.5">
                      <input 
                        type="color" 
                        value={customTheme.totalBg} 
                        onChange={(e) => {
                          const val = e.target.value;
                          setCustomTheme(prev => ({
                            ...prev,
                            totalBg: val,
                            totalText: isLightColor(val) ? "#7c2d12" : "#ffffff"
                          }));
                        }}
                        className="w-6 h-6 rounded cursor-pointer border border-slate-300 p-0"
                      />
                      <input 
                        type="text" 
                        value={customTheme.totalBg}
                        onChange={(e) => {
                          const val = e.target.value;
                          if (val.startsWith("#") && val.length <= 7) {
                            setCustomTheme(prev => ({
                              ...prev,
                              totalBg: val,
                              totalText: isLightColor(val) ? "#7c2d12" : "#ffffff"
                            }));
                          }
                        }}
                        className="bg-slate-50 border border-slate-250 rounded px-2 py-1 text-[11px] font-mono text-slate-700 w-20"
                      />
                    </div>
                  </div>

                  <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider block font-mono">Chữ Tổng cộng:</label>
                    <div className="flex items-center gap-1.5">
                      <input 
                        type="color" 
                        value={customTheme.totalText} 
                        onChange={(e) => setCustomTheme(prev => ({ ...prev, totalText: e.target.value }))}
                        className="w-6 h-6 rounded cursor-pointer border border-slate-300 p-0"
                      />
                      <input 
                        type="text" 
                        value={customTheme.totalText}
                        onChange={(e) => {
                          const val = e.target.value;
                          if (val.startsWith("#") && val.length <= 7) {
                            setCustomTheme(prev => ({ ...prev, totalText: val }));
                          }
                        }}
                        className="bg-slate-50 border border-slate-250 rounded px-2 py-1 text-[11px] font-mono text-slate-700 w-20"
                      />
                    </div>
                  </div>
                </div>
              </div>

            </div>
          </div>
        )}

        {/* COLLAPSIBLE FIELD CONFIGURATION PANEL */}
        {showConfigPanel && (
          <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 space-y-4 animate-fade-in font-sans">
            <div className="flex items-center gap-1.5 text-xs font-bold text-indigo-800 uppercase tracking-wider">
              <Info className="w-4 h-4 text-indigo-600" />
              THIẾT LẬP CÁC CỘT CHỨA THÔNG TIN ĐỊA BÀN TRONG FILE CỦA BẠN
            </div>
            <p className="text-[11px] text-slate-500">
              Hệ thống đã tự động nhận diện các cột tương ứng. Nếu chưa đúng ý, vui lòng lựa chọn lại đúng tên cột trong file dữ liệu của bạn ở bên dưới:
            </p>

            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-4">
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-600 uppercase tracking-wider block font-mono">Cột Tỉnh/Thành:</label>
                <select
                  value={configProvCol}
                  onChange={(e) => setConfigProvCol(e.target.value)}
                  className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-medium text-slate-800 w-full focus:outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer"
                >
                  <option value="">-- Không có / Bỏ qua --</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-600 uppercase tracking-wider block font-mono">Cột Huyện/Quận:</label>
                <select
                  value={configDistCol}
                  onChange={(e) => setConfigDistCol(e.target.value)}
                  className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-medium text-slate-800 w-full focus:outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer"
                >
                  <option value="">-- Không có / Bỏ qua --</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-600 uppercase tracking-wider block font-mono">Cột Tên Xã (Bắt buộc):</label>
                <select
                  value={configCommCol}
                  onChange={(e) => setConfigCommCol(e.target.value)}
                  className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-medium text-slate-800 w-full focus:outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer"
                >
                  <option value="">-- Chọn cột xã --</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-600 uppercase tracking-wider block font-mono">Cột Mã Xã:</label>
                <select
                  value={configCommCodeCol}
                  onChange={(e) => setConfigCommCodeCol(e.target.value)}
                  className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-medium text-slate-800 w-full focus:outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer"
                >
                  <option value="">-- Tự động / Không có --</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-600 uppercase tracking-wider block font-mono">Cột Loại Phiếu:</label>
                <select
                  value={configSheetCol}
                  onChange={(e) => setConfigSheetCol(e.target.value)}
                  className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs font-medium text-slate-800 w-full focus:outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer"
                >
                  <option value="">-- Không có / Bỏ qua --</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            </div>
          </div>
        )}

        {/* SELECT FILTER MENUS & RUN BUTTON (CHÍNH XÁC NHƯ TRONG ẢNH CHỤP MÀN HÌNH CỦA NGƯỜI DÙNG) */}
        <div className="bg-sky-50/55 p-5 rounded-2xl border border-sky-100 flex flex-col xl:flex-row items-end gap-4 font-sans shadow-inner">
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4 w-full text-slate-800">
            {/* Tỉnh/Thành phố selection */}
            <div className="space-y-1.5">
              <label className="text-[11px] font-bold text-sky-900 block font-mono">Tỉnh/Thành phố:</label>
              <select
                value={selectedProv}
                onChange={(e) => {
                  setSelectedProv(e.target.value);
                  setSelectedDist("Tất cả"); // Reset Huyện khi đổi Tỉnh
                }}
                className="bg-white border border-sky-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-sky-500 cursor-pointer w-full shadow-sm"
              >
                <option value="Tất cả">-- Tất cả Tỉnh/Thành --</option>
                {provincesList.map(p => <option key={p} value={p}>{p}</option>)}
              </select>
            </div>

            {/* Huyện/Quận selection */}
            <div className="space-y-1.5">
              <label className="text-[11px] font-bold text-sky-900 block font-mono">Huyện/Quận:</label>
              <select
                value={selectedDist}
                onChange={(e) => setSelectedDist(e.target.value)}
                className="bg-white border border-sky-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-sky-500 cursor-pointer w-full shadow-sm"
              >
                <option value="Tất cả">-- Tất cả Huyện/Quận --</option>
                {districtsList.map(d => <option key={d} value={d}>{d}</option>)}
              </select>
            </div>

            {/* Loại phiếu selection */}
            <div className="space-y-1.5">
              <label className="text-[11px] font-bold text-sky-900 block font-mono">Loại phiếu:</label>
              <select
                value={selectedSheet}
                onChange={(e) => setSelectedSheet(e.target.value)}
                className="bg-white border border-sky-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-sky-500 cursor-pointer w-full shadow-sm"
              >
                <option value="Tất cả">-- Tất cả loại phiếu --</option>
                {sheetsList.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>

            {/* Chỉ tiêu selection */}
            <div className="space-y-1.5">
              <label className="text-[11px] font-bold text-sky-900 block font-mono">Chỉ tiêu khảo sát:</label>
              <select
                value={configIndicatorCol}
                onChange={(e) => setConfigIndicatorCol(e.target.value)}
                className="bg-white border border-sky-200 rounded-xl px-3 py-2 text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-sky-500 cursor-pointer w-full shadow-sm"
              >
                <option value="">-- Chọn chỉ tiêu cần đếm --</option>
                {columns.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
          </div>

          {/* RUN ACTION BUTTON ("Xem KQ") */}
          <button
            onClick={handleRunCommuneAnalysis}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl cursor-pointer shadow-md transition-all active:scale-95 flex items-center gap-2 self-stretch xl:self-auto justify-center shrink-0 border-0"
          >
            <Play className="w-3.5 h-3.5 fill-current" /> Xem KQ
          </button>
        </div>

        {/* RESULTS TABLE AND ACTIONS */}
        {hasRunAnalysis && (
          <div className="space-y-4 animate-fade-in font-sans text-slate-800">
            
            {/* SUB BUTTONS ROW: EXCEL EXPORT & TABLE SEARCH */}
            <div className="flex flex-col sm:flex-row justify-between sm:items-center gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200">
              <div className="flex items-center gap-2">
                <button
                  onClick={handleExportCommuneExcel}
                  className="bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 font-bold text-xs py-1.5 px-3.5 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer shadow-xs"
                >
                  <Download className="w-3.5 h-3.5" /> Tải dữ liệu Excel (.xlsx)
                </button>
                <span className="text-[10px] text-slate-500 font-medium hidden sm:inline">
                  Nhấp đúp chuột vào bất kỳ ô số liệu nào trên bảng để <strong>chỉnh sửa số lại</strong> tùy ý.
                </span>
              </div>

              {/* SEARCH FILTER */}
              <div className="relative">
                <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none">
                  <Search className="h-3.5 w-3.5 text-slate-400" />
                </span>
                <input
                  type="text"
                  placeholder="Tìm kiếm theo tên xã..."
                  value={communeSearchTerm}
                  onChange={(e) => setCommuneSearchTerm(e.target.value)}
                  className="bg-white border border-slate-300 rounded-lg pl-9 pr-3 py-1.5 text-xs text-slate-800 placeholder-slate-450 focus:outline-none focus:ring-1 focus:ring-indigo-500 w-52 font-medium"
                />
              </div>
            </div>

            {/* THE MAIN TABLE CORE (MATCHES THE IMAGE LAYOUT EXACTLY) */}
            <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white shadow-md max-h-[600px]" style={{ borderColor: customTheme.borderColor }}>
              <table className="w-full text-left text-xs border-collapse" style={{ borderColor: customTheme.borderColor }}>
                <thead>
                  <tr 
                    className="border-b text-slate-700 font-bold text-[11px] uppercase tracking-wider sticky top-0 z-10 select-none"
                    style={{ backgroundColor: customTheme.headerBg, color: customTheme.headerText, borderColor: customTheme.borderColor }}
                  >
                    <th className="p-3 w-14 text-center">STT</th>
                    <th className="p-3 w-28 text-center">Mã xã</th>
                    <th className="p-3">Tên</th>
                    <th className="p-3 w-36 text-center">Tổng số</th>
                    <th className="p-3 w-36 text-center">Có</th>
                    <th className="p-3 w-36 text-center">Không</th>
                  </tr>
                </thead>

                <tbody className="divide-y text-slate-800 text-[11.5px]" style={{ borderColor: customTheme.borderColor }}>
                  {filteredCommuneRows.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="text-center py-8 text-slate-500">
                        Không tìm thấy xã nào phù hợp với từ khóa.
                      </td>
                    </tr>
                  ) : (
                    filteredCommuneRows.map((row, idx) => {
                      const stt = row.isTotalRow ? "" : communeRows.filter(r => !r.isTotalRow).indexOf(row) + 1;
                      const isTotal = row.isTotalRow;
                      
                      const rowStyle = isTotal 
                        ? { backgroundColor: customTheme.totalBg, color: customTheme.totalText, fontWeight: "bold" as const }
                        : { backgroundColor: idx % 2 === 1 ? customTheme.rowOddBg : customTheme.rowEvenBg };

                      return (
                        <tr
                          key={row.tenXa + "_" + row.maXa}
                          className={`hover:opacity-90 transition-opacity ${
                            isTotal 
                              ? "font-black border-b-2 shadow-xs sticky top-[31px] z-10" 
                              : ""
                          }`}
                          style={rowStyle}
                        >
                          {/* STT */}
                          <td className="p-3 text-center font-bold font-mono" style={{ opacity: 0.7 }}>
                            {stt}
                          </td>

                          {/* Mã xã */}
                          <td className="p-3 text-center font-bold font-mono" style={{ opacity: 0.8 }}>
                            {row.maXa}
                          </td>

                          {/* Tên Xã / Huyện lỵ */}
                          <td className={`p-3 font-semibold ${isTotal ? "text-[12.5px]" : ""}`}>
                            {row.tenXa}
                          </td>

                          {/* Tổng số (Editable input) */}
                          <td className="p-2 text-center">
                            <input
                              type="text"
                              value={row.tongSo}
                              onChange={(e) => handleCellChangeCommune(communeRows.indexOf(row), "tongSo", e.target.value)}
                              className="w-full bg-transparent text-center font-bold font-mono text-[12px] outline-none rounded py-1 border-0 focus:bg-amber-100/30 focus:ring-1 focus:ring-amber-500"
                              style={{ color: isTotal ? customTheme.totalText : "inherit" }}
                              title="Nhấp chuột để sửa tổng số lượng"
                            />
                          </td>

                          {/* Có (Editable input) */}
                          <td className="p-2 text-center">
                            <input
                              type="text"
                              value={row.co || ""}
                              onChange={(e) => handleCellChangeCommune(communeRows.indexOf(row), "co", e.target.value)}
                              className={`w-full bg-transparent text-center font-bold font-mono text-[12px] outline-none rounded py-1 border-0 focus:bg-amber-100/30 focus:ring-1 focus:ring-amber-500 ${
                                isTotal ? "text-[13px]" : "text-indigo-600 hover:text-indigo-800"
                              }`}
                              style={{ color: isTotal ? customTheme.totalText : "" }}
                              placeholder="0"
                              title="Nhấp chuột để sửa số 'Có'"
                            />
                          </td>

                          {/* Không (Editable input) */}
                          <td className="p-2 text-center">
                            <input
                              type="text"
                              value={row.khong || ""}
                              onChange={(e) => handleCellChangeCommune(communeRows.indexOf(row), "khong", e.target.value)}
                              className={`w-full bg-transparent text-center font-bold font-mono text-[12px] outline-none rounded py-1 border-0 focus:bg-amber-100/30 focus:ring-1 focus:ring-amber-500 ${
                                isTotal ? "text-[13px]" : "text-teal-600 hover:text-teal-800"
                              }`}
                              style={{ color: isTotal ? customTheme.totalText : "" }}
                              placeholder="0"
                              title="Nhấp chuột để sửa số 'Không'"
                            />
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

    </div>
  );
});

FrequencyAnalysis.displayName = "FrequencyAnalysis";
