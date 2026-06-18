import React, { useState } from 'react';
import * as XLSX from 'xlsx';

interface FilterOptions {
  includeStateOwned: boolean;
  includeSmallCommunes: boolean;
  smallCommuneThreshold: number;
  includeRevenueSample: boolean;
  revenueThreshold: number;
  prioritizeExistingSample: boolean;
  existingSampleFile: File | null;
  existingSampleKeyCol: string;
  backupCount: number;
}

export default function SampleSelector() {
  const [activeBlock, setActiveBlock] = useState<'enterprise' | 'individual'>('enterprise');

  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [rawData, setRawData] = useState<any[][]>([]);
  const [headers, setHeaders] = useState<string[]>([]);

  const [mapping, setMapping] = useState({
    maNganh: '',
    doanhThu: '',
    loaiHinh: '',
    tenDN: '',
    xaPhuong: '',
  });

  const [filterOptions, setFilterOptions] = useState<FilterOptions>({
    includeStateOwned: true,
    includeSmallCommunes: false,
    smallCommuneThreshold: 5,
    includeRevenueSample: false,
    revenueThreshold: 75,
    prioritizeExistingSample: false,
    existingSampleFile: null,
    existingSampleKeyCol: '',
    backupCount: 2,
  });

  const [filteredData, setFilteredData] = useState<any[]>([]);
  const [showResults, setShowResults] = useState(false);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setFile(f);
    const buffer = await f.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    setSheetNames(wb.SheetNames);
    if (wb.SheetNames.length > 0) {
      const sheet = wb.SheetNames[0];
      setSelectedSheet(sheet);
      const ws = wb.Sheets[sheet];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
      setRawData(data);
      if (data.length > 0) setHeaders(data[0].map((h: any) => String(h)));
    }
    setShowResults(false);
  };

  const handleSheetChange = async (sheet: string) => {
    if (!file) return;
    setSelectedSheet(sheet);
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const ws = wb.Sheets[sheet];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
    setRawData(data);
    if (data.length > 0) setHeaders(data[0].map((h: any) => String(h)));
    setShowResults(false);
  };

  const handleExistingSampleUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) setFilterOptions(prev => ({ ...prev, existingSampleFile: f }));
  };

  const isIndustry = (ma: string) => {
    const prefix = ma.substring(0, 2).toUpperCase();
    return /^\d{2}$/.test(prefix) && parseInt(prefix) >= 10 && parseInt(prefix) <= 33;
  };

  const applyFilters = async () => {
    if (rawData.length < 2) return;

    const dataRows = rawData.slice(1).map(row => {
      const obj: any = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });

    const getVal = (row: any, field: keyof typeof mapping) => {
      const colName = mapping[field];
      return colName ? row[colName] : undefined;
    };

    const normalized = dataRows.map(row => ({
      ...row,
      _maNganh: String(getVal(row, 'maNganh') || '').trim(),
      _doanhThu: parseFloat(getVal(row, 'doanhThu') || '0') || 0,
      _loaiHinh: String(getVal(row, 'loaiHinh') || '').trim().toLowerCase(),
      _tenDN: String(getVal(row, 'tenDN') || '').trim(),
      _xaPhuong: String(getVal(row, 'xaPhuong') || '').trim(),
      _originalRow: row,
    }));

    let existingSampleKeys: Set<string> = new Set();
    if (filterOptions.prioritizeExistingSample && filterOptions.existingSampleFile && filterOptions.existingSampleKeyCol) {
      try {
        const buffer = await filterOptions.existingSampleFile.arrayBuffer();
        const wb = XLSX.read(buffer, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const oldData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
        if (oldData.length > 1) {
          const oldHeaders = oldData[0].map((h: any) => String(h));
          const oldRows = oldData.slice(1).map(row => {
            const obj: any = {};
            oldHeaders.forEach((h, i) => { obj[h] = row[i]; });
            return obj;
          });
          const keyCol = filterOptions.existingSampleKeyCol;
          oldRows.forEach(r => {
            const val = String(r[keyCol] || '').trim();
            if (val) existingSampleKeys.add(val);
          });
        }
      } catch (err) {
        console.warn('Không đọc được file mẫu cũ:', err);
      }
    }

    let resultWithTags: any[] = [];

    if (activeBlock === 'enterprise') {
      if (filterOptions.includeStateOwned) {
        const filtered = normalized.filter(r =>
          (r._loaiHinh === 'nhà nước' || r._loaiHinh === 'nn') && isIndustry(r._maNganh)
        );
        filtered.forEach(r => resultWithTags.push({ ...r._originalRow, Loai: 'Chính' }));
      }

      if (filterOptions.includeSmallCommunes) {
        const communeCount: Record<string, number> = {};
        normalized.forEach(r => {
          if (r._xaPhuong && isIndustry(r._maNganh)) {
            communeCount[r._xaPhuong] = (communeCount[r._xaPhuong] || 0) + 1;
          }
        });
        const smallXas = Object.entries(communeCount)
          .filter(([_, count]) => count < filterOptions.smallCommuneThreshold)
          .map(([xa]) => xa);
        const filtered = normalized.filter(r =>
          smallXas.includes(r._xaPhuong) && isIndustry(r._maNganh)
        );
        filtered.forEach(r => resultWithTags.push({ ...r._originalRow, Loai: 'Chính' }));
      }

      if (filterOptions.includeRevenueSample) {
        const nonState = normalized.filter(r =>
          r._loaiHinh !== 'nhà nước' && r._loaiHinh !== 'nn' && isIndustry(r._maNganh)
        );
        const groups: Record<string, any[]> = {};
        nonState.forEach(r => {
          const cap2 = r._maNganh.substring(0, 2);
          if (!groups[cap2]) groups[cap2] = [];
          groups[cap2].push(r);
        });
        for (const [cap2, items] of Object.entries(groups)) {
          const sorted = items.sort((a, b) => {
            const aKey = String(getVal(a._originalRow, 'tenDN') || '').trim();
            const bKey = String(getVal(b._originalRow, 'tenDN') || '').trim();
            const aPriority = existingSampleKeys.has(aKey) ? 1 : 0;
            const bPriority = existingSampleKeys.has(bKey) ? 1 : 0;
            if (aPriority !== bPriority) return bPriority - aPriority;
            return b._doanhThu - a._doanhThu;
          });
          const totalRevenue = sorted.reduce((sum, r) => sum + r._doanhThu, 0);
          if (totalRevenue === 0) continue;
          let cumulative = 0;
          let cutIndex = sorted.length;
          for (let i = 0; i < sorted.length; i++) {
            cumulative += sorted[i]._doanhThu;
            if ((cumulative / totalRevenue) * 100 >= filterOptions.revenueThreshold) {
              cutIndex = i + 1;
              break;
            }
          }
          for (let i = 0; i < cutIndex; i++) {
            resultWithTags.push({ ...sorted[i]._originalRow, Loai: 'Chính' });
          }
          const backupStart = cutIndex;
          const backupEnd = Math.min(backupStart + filterOptions.backupCount, sorted.length);
          for (let i = backupStart; i < backupEnd; i++) {
            resultWithTags.push({ ...sorted[i]._originalRow, Loai: 'Dự phòng' });
          }
        }
      }
    } else {
      const individualData = normalized.filter(r => isIndustry(r._maNganh));
      const grouped: Record<string, any[]> = {};
      individualData.forEach(r => {
        const cap2 = r._maNganh.substring(0, 2);
        const key = `${r._xaPhuong}|${cap2}`;
        if (!grouped[key]) grouped[key] = [];
        grouped[key].push(r);
      });
      for (const [key, items] of Object.entries(grouped)) {
        const [xa, cap2] = key.split('|');
        const count = items.length;
        let sampleSize = 0;
        if (count <= 5) sampleSize = count;
        else if (count <= 100) sampleSize = 5;
        else if (count <= 1000) sampleSize = 8;
        else sampleSize = Math.ceil(count * 0.01);
        if (sampleSize === 0) continue;
        const sorted = items.sort((a, b) => b._doanhThu - a._doanhThu);
        for (let i = 0; i < Math.min(sampleSize, sorted.length); i++) {
          resultWithTags.push({ ...sorted[i]._originalRow, Loai: 'Chính' });
        }
        const backupStart = sampleSize;
        const backupEnd = Math.min(backupStart + filterOptions.backupCount, sorted.length);
        for (let i = backupStart; i < backupEnd; i++) {
          resultWithTags.push({ ...sorted[i]._originalRow, Loai: 'Dự phòng' });
        }
      }
      const maxPerXa = 100;
      const xaCount: Record<string, number> = {};
      const finalResult = [];
      for (const item of resultWithTags) {
        const xa = String(getVal(item, 'xaPhuong') || '').trim();
        if (item.Loai === 'Chính') {
          xaCount[xa] = (xaCount[xa] || 0) + 1;
          finalResult.push(item);
        }
      }
      for (const item of resultWithTags) {
        if (item.Loai === 'Dự phòng') {
          const xa = String(getVal(item, 'xaPhuong') || '').trim();
          if ((xaCount[xa] || 0) < maxPerXa) {
            xaCount[xa] = (xaCount[xa] || 0) + 1;
            finalResult.push(item);
          }
        }
      }
      resultWithTags = finalResult;
    }

    const uniqueResult = Array.from(new Set(resultWithTags.map(r => JSON.stringify(r)))).map(s => JSON.parse(s));
    setFilteredData(uniqueResult);
    setShowResults(true);
  };

  const exportResult = () => {
    if (filteredData.length === 0) return;
    const ws = XLSX.utils.json_to_sheet(filteredData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'KetQua');
    XLSX.writeFile(wb, 'MauChonLoc.xlsx');
  };

  return (
    <div className="bg-[#1f2937]/80 text-gray-200 rounded-2xl p-6 space-y-6">
      <h2 className="text-2xl font-bold text-purple-300">🎯 Chọn mẫu điều tra</h2>

      <div className="flex gap-4">
        <button onClick={() => setActiveBlock('enterprise')}
          className={`px-4 py-2 rounded-lg font-semibold text-sm transition ${activeBlock === 'enterprise' ? 'bg-purple-600 text-white' : 'bg-gray-700 text-gray-300'}`}>
          🏭 Khối Doanh nghiệp
        </button>
        <button onClick={() => setActiveBlock('individual')}
          className={`px-4 py-2 rounded-lg font-semibold text-sm transition ${activeBlock === 'individual' ? 'bg-green-600 text-white' : 'bg-gray-700 text-gray-300'}`}>
          🏪 Khối Cá thể
        </button>
      </div>

      <div className="border border-gray-600 p-4 rounded-lg bg-[#111827]/80">
        <label className="font-semibold">📁 File dữ liệu (Doanh nghiệp/Cá thể):</label>
        <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload}
          className="mt-2 block w-full text-sm text-gray-300 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:bg-purple-500 file:text-white file:cursor-pointer" />
      </div>

      {sheetNames.length > 0 && (
        <div className="flex items-center gap-3">
          <label className="font-medium">Sheet:</label>
          <select value={selectedSheet} onChange={(e) => handleSheetChange(e.target.value)}
            className="bg-[#111827] border border-gray-600 rounded px-3 py-1 text-sm text-white">
            {sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <span className="text-xs text-gray-400">({rawData.length - 1} dòng dữ liệu)</span>
        </div>
      )}

      {headers.length > 0 && (
        <div className="bg-[#111827]/80 p-4 rounded-lg border border-gray-600">
          <h3 className="font-semibold mb-3">📋 Chọn cột tương ứng</h3>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
            {[
              { key: 'maNganh', label: 'Mã ngành' },
              { key: 'doanhThu', label: 'Doanh thu' },
              { key: 'loaiHinh', label: 'Loại hình (Nhà nước/Ngoài NN/Cá thể)' },
              { key: 'tenDN', label: 'Tên doanh nghiệp/Cơ sở' },
              { key: 'xaPhuong', label: 'Xã/Phường' },
            ].map(item => (
              <div key={item.key} className="flex flex-col gap-1">
                <label className="text-xs font-medium text-gray-400">{item.label}</label>
                <select
                  value={(mapping as any)[item.key] || ''}
                  onChange={(e) => setMapping(prev => ({ ...prev, [item.key]: e.target.value }))}
                  className="bg-[#0f172a] border border-gray-600 rounded px-2 py-1 text-xs text-white"
                >
                  <option value="">-- Chọn --</option>
                  {headers.map((h, i) => <option key={i} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
        </div>
      )}

      {rawData.length > 0 && (
        <div className="border border-gray-600 p-4 rounded-lg space-y-4 bg-[#111827]/80">
          <h3 className="font-semibold text-lg text-amber-400">⚙️ Điều kiện lọc ({activeBlock === 'enterprise' ? 'Doanh nghiệp' : 'Cá thể'})</h3>

          {activeBlock === 'enterprise' ? (
            <>
              <label className="flex items-center gap-2 cursor-pointer">
                <input type="checkbox" checked={filterOptions.includeStateOwned}
                  onChange={e => setFilterOptions(prev => ({ ...prev, includeStateOwned: e.target.checked }))} />
                <span>Lấy toàn bộ <strong>Doanh nghiệp Nhà nước</strong></span>
              </label>

              <div className="flex items-center gap-4 flex-wrap">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input type="checkbox" checked={filterOptions.includeSmallCommunes}
                    onChange={e => setFilterOptions(prev => ({ ...prev, includeSmallCommunes: e.target.checked }))} />
                  <span>Lấy <strong>DN công nghiệp cấp 2</strong> ở xã có số DN &lt; </span>
                </label>
                <input type="number" value={filterOptions.smallCommuneThreshold}
                  onChange={e => setFilterOptions(prev => ({ ...prev, smallCommuneThreshold: parseInt(e.target.value) || 5 }))}
                  className="w-20 bg-[#0f172a] border border-gray-600 rounded px-2 py-1 text-sm text-white" disabled={!filterOptions.includeSmallCommunes} />
              </div>

              <div className="space-y-2">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input type="checkbox" checked={filterOptions.includeRevenueSample}
                    onChange={e => setFilterOptions(prev => ({ ...prev, includeRevenueSample: e.target.checked }))} />
                  <span>Chọn mẫu <strong>DN ngoài Nhà nước</strong> theo tỷ trọng doanh thu cộng dồn</span>
                </label>
                <div className="flex items-center gap-2 ml-6">
                  <span>Ngưỡng (%): </span>
                  <input type="number" value={filterOptions.revenueThreshold}
                    onChange={e => setFilterOptions(prev => ({ ...prev, revenueThreshold: parseFloat(e.target.value) || 75 }))}
                    className="w-20 bg-[#0f172a] border border-gray-600 rounded px-2 py-1 text-sm text-white" disabled={!filterOptions.includeRevenueSample} />
                  <span>%</span>
                </div>
              </div>

              <div className="space-y-2">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input type="checkbox" checked={filterOptions.prioritizeExistingSample}
                    onChange={e => setFilterOptions(prev => ({ ...prev, prioritizeExistingSample: e.target.checked }))} />
                  <span>Ưu tiên lấy DN có trong file mẫu tháng trước</span>
                </label>
                {filterOptions.prioritizeExistingSample && (
                  <div className="ml-6 space-y-2">
                    <div>
                      <label className="text-xs text-gray-400">Chọn cột định danh (mã số thuế, ID...):</label>
                      <select
                        value={filterOptions.existingSampleKeyCol}
                        onChange={e => setFilterOptions(prev => ({ ...prev, existingSampleKeyCol: e.target.value }))}
                        className="bg-[#0f172a] border border-gray-600 rounded px-2 py-1 text-xs text-white mt-1"
                      >
                        <option value="">-- Chọn cột --</option>
                        {headers.map(h => <option key={h} value={h}>{h}</option>)}
                      </select>
                    </div>
                    <input type="file" accept=".xlsx,.xls" onChange={handleExistingSampleUpload}
                      className="text-sm text-gray-300 file:py-1 file:px-3 file:rounded file:bg-gray-600 file:text-gray-200 file:cursor-pointer" />
                  </div>
                )}
              </div>
            </>
          ) : (
            <div className="text-sm text-gray-300 bg-gray-800/50 p-3 rounded-lg">
              <p>✅ <strong>Khối Cá thể:</strong> Tự động xác định cỡ mẫu theo ngành cấp 2 (1-5: chọn hết; 6-100: 5; 101-1000: 8; &gt;1000: 1%) và chọn các cơ sở có doanh thu cao nhất.</p>
              <p className="text-xs text-gray-400 mt-1">Giới hạn tối đa 100 cơ sở/xã. Chỉ lấy ngành công nghiệp (mã 10-33).</p>
            </div>
          )}

          <div className="flex items-center gap-4">
            <label className="text-sm">Số lượng dự phòng mỗi nhóm:</label>
            <input type="number" value={filterOptions.backupCount}
              onChange={e => setFilterOptions(prev => ({ ...prev, backupCount: parseInt(e.target.value) || 0 }))}
              className="w-20 bg-[#0f172a] border border-gray-600 rounded px-2 py-1 text-sm text-white" min={0} />
          </div>

          <button onClick={applyFilters}
            className="px-6 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg font-semibold transition">
            🔍 Chạy lọc
          </button>
        </div>
      )}

      {showResults && (
        <div className="border border-gray-600 p-4 rounded-lg bg-[#111827]/80">
          <div className="flex justify-between items-center mb-3">
            <h3 className="font-semibold text-lg text-emerald-400">📊 Kết quả lọc ({filteredData.length} dòng)</h3>
            <button onClick={exportResult}
              className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg transition">
              💾 Xuất Excel
            </button>
          </div>
          {filteredData.length > 0 ? (
            <div className="overflow-auto max-h-96">
              <table className="min-w-full text-xs border border-gray-600">
                <thead>
                  <tr className="bg-gray-700">
                    {Object.keys(filteredData[0]).map((key) => (
                      <th key={key} className="p-2 border border-gray-600 text-left">{key}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((row, idx) => (
                    <tr key={idx} className="border-t border-gray-600">
                      {Object.values(row).map((val, i) => (
                        <td key={i} className="p-2 border border-gray-600">{String(val)}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="text-gray-400">Không có dữ liệu phù hợp.</p>
          )}
        </div>
      )}
    </div>
  );
}