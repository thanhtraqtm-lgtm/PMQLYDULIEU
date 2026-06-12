import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { 
  Home, FileSpreadsheet, GitMerge, Combine, Scissors, BarChart3, PieChart,
  CheckSquare, Download, Loader2, FileUp, AlertTriangle, CheckCircle2, Brain,
  Layers, ArrowRight, ArrowRightLeft, Database, Search, Plus, Lock, KeyRound, LogOut,
  ChevronLeft, ChevronRight
} from "lucide-react";
import { 
  normalizeSectorCode, getSectorLevel, getParentSectorCode,
  lookupSectorNameWithFallback, clearAllSectorsInVSIC, loadSectorsIntoVSIC
} from "./data/vsic";
import SectorRevenueChart from "./components/sectorRevenueChart";
import VsicCatalogExplorer from "./components/vsicCatalogExplorer";
import DescriptorMatchScanner from "./components/descriptorMatchScanner";
import { BeautifulReportTable } from "./components/BeautifulReportTable";

// Helper: component bảng dữ liệu đẹp
const DataPreviewTable = ({ data, columns, title }: { data: any[]; columns: string[]; title?: string }) => {
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(1);
  const pageSize = 50;
  const filtered = useMemo(() => {
    if (!search) return data;
    const term = search.toLowerCase();
    return data.filter(row => Object.values(row).some(v => String(v).toLowerCase().includes(term)));
  }, [data, search]);
  const totalPages = Math.ceil(filtered.length / pageSize);
  const paginated = filtered.slice((page-1)*pageSize, page*pageSize);
  if (!data.length) return null;
  return (
    <div className="mt-6 bg-[#1f2937]/80 rounded-2xl border border-gray-800 p-4">
      <div className="flex justify-between items-center mb-3">
        <div className="flex items-center gap-2"><Database className="w-4 h-4 text-cyan-400"/><span className="font-bold">{title||"Kết quả"}</span><span className="text-xs text-gray-400">{data.length} dòng</span></div>
        <div className="relative"><Search className="absolute left-2 top-1.5 w-3.5 h-3.5 text-gray-500"/><input className="pl-7 pr-2 py-1 bg-[#111827] border border-gray-700 rounded-lg text-xs" placeholder="Tìm..." value={search} onChange={e=>{setSearch(e.target.value); setPage(1);}}/></div>
      </div>
      <div className="overflow-x-auto"><table className="w-full text-xs"><thead className="bg-[#0f172a]"><tr>{columns.map(c=><th key={c} className="p-2 text-center">{c}</th>)}</thead><tbody>{paginated.map((row,i)=>(
        <tr key={i} className="border-t border-gray-800">{columns.map(c=><td key={c} className="p-2 text-center truncate max-w-[180px]">{row[c]??""}</td>)}</tr>
      ))}</tbody></table></div>
      <div className="flex justify-between mt-3 text-xs"><span>Trang {page}/{totalPages||1}</span><div className="flex gap-2"><button disabled={page===1} onClick={()=>setPage(p=>p-1)}><ChevronLeft className="w-4 h-4"/></button><button disabled={page===totalPages} onClick={()=>setPage(p=>p+1)}><ChevronRight className="w-4 h-4"/></button></div></div>
    </div>
  );
};

interface ColumnMapping { mota: string; manganh: string; xa: string; doanhthu: string; laodong: string; idCol: string; }
      export default function App() {
  // ==================== AUTH ====================
  const [isAuthorized, setIsAuthorized] = useState(()=>localStorage.getItem("vsic_app_authorized")==="true");
  const [typedPwd, setTypedPwd] = useState("");
  const [pwdErr, setPwdErr] = useState("");
  const appPwd = "admin123";

  // ==================== CORE DATA ====================
  const [activeTab, setActiveTab] = useState("trangchu");
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [msg, setMsg] = useState("");
  const [mainData, setMainData] = useState<any[]>([]);
  const [rawData, setRawData] = useState<any[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [fileName, setFileName] = useState("");
  const [mapping, setMapping] = useState<ColumnMapping>({ mota:"", manganh:"", xa:"", doanhthu:"", laodong:"", idCol:"" });
  const [colConfigs, setColConfigs] = useState<{originalName:string; use:boolean; newName:string; role:""|"mota"|"manganh"|"xa"|"doanhthu"|"laodong"|"idCol"}[]>([]);
  const [userSectorMap, setUserSectorMap] = useState<Map<string,string>>(new Map());
  const [userSectorFile, setUserSectorFile] = useState("");
  const [expanded, setExpanded] = useState(true);
  const [reportType, setReportType] = useState<"flat"|"pivot">("pivot");
  const [quickRows, setQuickRows] = useState<any[]>([]);
  const [quickCols, setQuickCols] = useState<string[]>([]);
  const [quickLevel, setQuickLevel] = useState(1);
  const [searchTerm, setSearchTerm] = useState("");
  const [pageView, setPageView] = useState(1);
  const pageSize = 50;

  // ==================== FEATURE STATES ====================
  const [oldData, setOldData] = useState<any[]>([]); const [oldName, setOldName]=useState("");
  const [newData, setNewData]=useState<any[]>([]); const [newName,setNewName]=useState("");
  const [diffKey,setDiffKey]=useState("");
  const [leftData,setLeftData]=useState<any[]>([]); const [leftName,setLeftName]=useState("");
  const [rightData,setRightData]=useState<any[]>([]); const [rightName,setRightName]=useState("");
  const [leftKey,setLeftKey]=useState(""); const [rightKey,setRightKey]=useState("");
  const [splitCol,setSplitCol]=useState("");
  const [detectedSheets,setDetectedSheets]=useState<string[]>([]);
  const [selectedSheets,setSelectedSheets]=useState<string[]>([]);
  const [sheetKey,setSheetKey]=useState("");
  const [workbook,setWorkbook]=useState<any>(null);
  const [qrManganh, setQrManganh]=useState(""); const [qrXa, setQrXa]=useState("");
  const [qrDoanhthu, setQrDoanhthu]=useState(""); const [qrLaodong, setQrLaodong]=useState("");
  const [stdCol, setStdCol]=useState(""); const [stdMatch, setStdMatch]=useState({total:0,valid:0,invalid:0});
  const [compA, setCompA]=useState(""); const [compB, setCompB]=useState(""); const [compRule, setCompRule]=useState("normalize");
  const [calcName, setCalcName]=useState(""); const [calcCol1, setCalcCol1]=useState(""); const [calcCol2, setCalcCol2]=useState("");
  const [calcOp, setCalcOp]=useState<"+"|"-"|"*"|"/"|"concat">("+");
  const [calcType, setCalcType]=useState<"column"|"constant">("column");
  const [calcConst, setCalcConst]=useState(""); const [calcRound, setCalcRound]=useState<"none"|"int"|"1dec"|"2dec">("none");

  // Helper functions
  const sleep = (ms:number)=>new Promise(r=>setTimeout(r,ms));
  const parseCSV = (text:string)=>{
    const rows = text.split(/\r?\n/).filter(r=>r.trim());
    if(!rows.length) return [];
    const headers = rows[0].split(",").map(h=>h.trim());
    return rows.slice(1).map(row=>{
      const vals = row.split(",");
      const obj:any={};
      headers.forEach((h,i)=>obj[h]=vals[i]||"");
      return obj;
    });
  };
  const chunkProcess = async (arr:any[], size:number, fn:any, onProgress?:any)=>{
    const res=[]; const len=arr.length;
    for(let i=0;i<len;i+=size){
      const chunk=arr.slice(i,i+size);
      for(let j=0;j<chunk.length;j++) res.push(fn(chunk[j],i+j));
      if(onProgress) onProgress(Math.round(i/len*100));
      await sleep(0);
    }
    return res;
  };

  // ==================== CÁC HÀM XỬ LÝ CHÍNH ====================
  // (Upload danh mục, upload file, merge, compare, split, redefine, quick report, standardize, cross compare, append levels, calc column, export...)
  // Để tránh quá dài, tôi sẽ đặt các hàm này trong phần 3 – bạn sẽ copy tiếp.
  // Hãy chắc chắn bạn đã copy hết phần 2, sau đó copy phần 3 bên dưới.
    
