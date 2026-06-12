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
      // ==================== UPLOAD DANH MỤC NGÀNH ====================
  const handleUploadCatalog = async (e:React.ChangeEvent<HTMLInputElement>)=>{
    const file=e.target.files?.[0]; if(!file) return;
    setLoading(true); setMsg("Đang nạp danh mục...");
    const reader=new FileReader();
    reader.onload=async(ev)=>{
      try{
        let rows:any[];
        if(file.name.endsWith(".csv")) rows=parseCSV(ev.target?.result as string);
        else{
          const wb=XLSX.read(ev.target?.result,{type:"array",dense:true});
          const ws=wb.Sheets[wb.SheetNames[0]];
          rows=XLSX.utils.sheet_to_json(ws);
        }
        const first=rows[0];
        let codeCol=Object.keys(first).find(k=>/mã|ma|code/i.test(k))||Object.keys(first)[0];
        let nameCol=Object.keys(first).find(k=>/tên|ten|name/i.test(k))||Object.keys(first)[1]||codeCol;
        const newMap=new Map();
        rows.forEach(r=>{const c=String(r[codeCol]||"").trim(); const n=String(r[nameCol]||"").trim(); if(c&&n) newMap.set(c,n);});
        if(!newMap.size) throw new Error("Không tìm thấy cặp mã-tên");
        clearAllSectorsInVSIC();
        loadSectorsIntoVSIC(Object.fromEntries(newMap));
        localStorage.setItem("custom_vsic_data",JSON.stringify(Object.fromEntries(newMap)));
        localStorage.setItem("custom_vsic_filename",file.name);
        setUserSectorMap(newMap); setUserSectorFile(file.name);
        setMsg(`Đã nạp ${newMap.size} mã ngành`);
      }catch(e:any){ alert(e.message); } finally{ setLoading(false); }
    };
    if(file.name.endsWith(".csv")) reader.readAsText(file,"UTF-8");
    else reader.readAsArrayBuffer(file);
  };
  const handleClearCatalog = ()=>{
    clearAllSectorsInVSIC(); localStorage.removeItem("custom_vsic_data");
    setUserSectorMap(new Map()); setUserSectorFile("");
  };

  // Upload file chính
  const handleFileUpload = (e:React.ChangeEvent<HTMLInputElement>, type:"main"|"old"|"new"|"left"|"right")=>{
    const file=e.target.files?.[0]; if(!file) return;
    setLoading(true); setMsg(`Đang tải ${file.name}...`);
    const reader=new FileReader();
    reader.onload=(ev)=>{
      try{
        let data:any[];
        if(file.name.endsWith(".csv")) data=parseCSV(ev.target?.result as string);
        else{
          const wb=XLSX.read(ev.target?.result,{type:"array",dense:true});
          const ws=wb.Sheets[wb.SheetNames[0]];
          data=XLSX.utils.sheet_to_json(ws);
          if(type==="main"){ setWorkbook(wb); setDetectedSheets(wb.SheetNames); setSelectedSheets(wb.SheetNames); }
        }
        const cols=Object.keys(data[0]||{});
        if(type==="main"){
          setRawData(data); setMainData(data); setColumns(cols); setFileName(file.name);
          const init=cols.map(c=>({originalName:c, use:true, newName:c, role:""}));
          setColConfigs(init); setMapping({mota:"",manganh:"",xa:"",doanhthu:"",laodong:"",idCol:""});
          setActiveTab("xemdulieu");
        }else if(type==="old"){ setOldData(data); setOldName(file.name); }
        else if(type==="new"){ setNewData(data); setNewName(file.name); }
        else if(type==="left"){ setLeftData(data); setLeftName(file.name); }
        else if(type==="right"){ setRightData(data); setRightName(file.name); }
        setMsg(`Đã tải ${data.length} dòng`);
      }catch(e:any){ alert("Lỗi: "+e.message); } finally{ setLoading(false); }
    };
    if(file.name.endsWith(".csv")) reader.readAsText(file,"UTF-8");
    else reader.readAsArrayBuffer(file);
  };

  // Ghép sheets
  const handleMergeSheets = async()=>{
    if(!workbook || selectedSheets.length<2 || !sheetKey) {alert("Chọn sheets và cột khóa"); return;}
    setLoading(true);
    const merged=new Map(); const allCols=new Set<string>();
    for(const s of selectedSheets){
      const ws=workbook.Sheets[s];
      if(!ws) continue;
      const rows=XLSX.utils.sheet_to_json(ws);
      rows.forEach((row:any)=>{
        const key=String(row[sheetKey]||"").trim()||`_no_${Math.random()}`;
        merged.set(key, merged.has(key)?{...merged.get(key),...row}:{...row});
        Object.keys(row).forEach(k=>allCols.add(k));
      });
    }
    const list=Array.from(merged.values()); const newCols=Array.from(allCols);
    setRawData(list); setMainData(list); setColumns(newCols);
    setColConfigs(newCols.map(c=>({originalName:c,use:true,newName:c,role:""})));
    setLoading(false); setMsg("Ghép sheets thành công!");
  };

  // Ghép nối trái-phải
  const handleMerge = async()=>{
    if(!leftData.length||!rightData.length||!leftKey||!rightKey){alert("Thiếu dữ liệu hoặc khóa"); return;}
    setLoading(true);
    const rightMap=new Map(); rightData.forEach(r=>{const k=String(r[rightKey]||"").trim(); if(k) rightMap.set(k,r);});
    const merged=[];
    const rightCols=Object.keys(rightData[0]||{}).filter(c=>c!==rightKey);
    for(const l of leftData){
      const k=String(l[leftKey]||"").trim();
      const m=rightMap.get(k);
      const row={...l};
      rightCols.forEach(c=>{const newName=row.hasOwnProperty(c)?`${c}_Phai`:c; row[newName]=m?m[c]:"";});
      merged.push(row);
    }
    const cols=Object.keys(merged[0]||{});
    setMainData(merged); setRawData(merged); setColumns(cols);
    setColConfigs(cols.map(c=>({originalName:c,use:true,newName:c,role:""})));
    setLoading(false); setMsg(`Ghép nối thành công ${merged.length} dòng`);
  };

  // So sánh cũ-mới
  const handleCompare = async()=>{
    if(!oldData.length||!newData.length||!diffKey){alert("Thiếu dữ liệu hoặc khóa"); return;}
    setLoading(true);
    const oldMap=new Map(); oldData.forEach(r=>{const k=String(r[diffKey]||"").trim(); if(k) oldMap.set(k,r);});
    const newMap=new Map(); newData.forEach(r=>{const k=String(r[diffKey]||"").trim(); if(k) newMap.set(k,r);});
    const allKeys=Array.from(new Set([...oldMap.keys(),...newMap.keys()]));
    const oldCols=Object.keys(oldData[0]||{}); const newCols=Object.keys(newData[0]||{});
    const unionCols=Array.from(new Set([...oldCols,...newCols])).filter(c=>c!==diffKey);
    const res=[];
    for(const k of allKeys){
      const o=oldMap.get(k); const n=newMap.get(k);
      const row:any={[diffKey]:k};
      if(o&&!n){
        unionCols.forEach(c=>{row[`${c}_Cu`]=o[c]||""; row[`${c}_Moi`]="";});
        row["Trạng_thái"]="❌ Đã xóa";
      }else if(!o&&n){
        unionCols.forEach(c=>{row[`${c}_Cu`]=""; row[`${c}_Moi`]=n[c]||"";});
        row["Trạng_thái"]="✅ Mới thêm";
      }else{
        const changed=[];
        unionCols.forEach(c=>{
          const v1=String(o[c]??"").trim(); const v2=String(n[c]??"").trim();
          row[`${c}_Cu`]=o[c]??""; row[`${c}_Moi`]=n[c]??"";
          if(v1!==v2) changed.push(c);
        });
        row["Trạng_thái"]=changed.length?`⚠️ Thay đổi: ${changed.join(", ")}`:"💡 Không đổi";
      }
      res.push(row);
    }
    const cols=Object.keys(res[0]||{});
    setMainData(res); setRawData(res); setColumns(cols);
    setColConfigs(cols.map(c=>({originalName:c,use:true,newName:c,role:""})));
    setLoading(false); setMsg(`So sánh xong ${res.length} khóa`);
  };

  // Tách file theo cột
  const handleSplit = async()=>{
    if(!mainData.length||!splitCol){alert("Chọn cột tách"); return;}
    setLoading(true);
    const groups=new Map();
    mainData.forEach(row=>{
      let val=String(row[splitCol]||"Khác").trim().replace(/[\\/:*?"<>|]/g,"_");
      if(!groups.has(val)) groups.set(val,[]);
      groups.get(val).push(row);
    });
    const zip=new JSZip(); let i=0;
    for(const [key,rows] of groups.entries()){
      const ws=XLSX.utils.json_to_sheet(rows);
      const wb=XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb,ws,"Data");
      zip.file(`Tach_${key}.xlsx`, XLSX.write(wb,{bookType:"xlsx",type:"array"}));
      setProgress(Math.round(++i/groups.size*100)); await sleep(10);
    }
    const blob=await zip.generateAsync({type:"blob"});
    const link=document.createElement("a"); link.href=URL.createObjectURL(blob); link.download=`Tach_${splitCol}.zip`; link.click();
    setLoading(false); setMsg(`Đã tách ${groups.size} file`);
  };

  // Định nghĩa lại cột (lọc + đổi tên)
  const handleRedefine = async()=>{
    const active=colConfigs.filter(c=>c.use && c.newName.trim());
    if(!active.length){alert("Chọn ít nhất một cột"); return;}
    setLoading(true);
    const newRows=await chunkProcess(rawData,10000,row=>{
      const newRow:any={};
      active.forEach(cfg=>newRow[cfg.newName.trim()]=row[cfg.originalName]??"");
      return newRow;
    });
    const newCols=Object.keys(newRows[0]||{});
    const newMap:ColumnMapping={mota:"",manganh:"",xa:"",doanhthu:"",laodong:"",idCol:""};
    Object.keys(newMap).forEach(rk=>{
      const old= mapping[rk as keyof ColumnMapping];
      if(old){
        const found=colConfigs.find(c=>c.originalName===old && c.use);
        if(found) newMap[rk as keyof ColumnMapping]=found.newName.trim();
      }
    });
    setMainData(newRows); setColumns(newCols); setMapping(newMap);
    setLoading(false); setMsg("Đã tái cấu trúc bảng");
  };

  // Báo cáo nhanh theo cấp ngành
  const handleQuickReport = async(level:number)=>{
    if(!mainData.length){alert("Chưa có dữ liệu"); return;}
    const targetManganh=qrManganh||mapping.manganh;
    const targetXa=qrXa||mapping.xa;
    if(!targetManganh||!targetXa){alert("Chọn cột mã ngành và xã"); return;}
    setLoading(true);
    const processed=await chunkProcess(mainData,5000,row=>{
      const mng=normalizeSectorCode(row[targetManganh]);
      let label="";
      if(level===2){
        const code=mng.slice(0,2); const name=userSectorMap.get(code)||"";
        label=code?`${code} - ${name}`:"Chưa xác định";
      }else{
        let code="";
        if(mng){ if(/^[A-Z]$/.test(mng)) code=mng.toUpperCase(); else code=getParentSectorCode(mng)||""; }
        const name=userSectorMap.get(code)||"";
        label=code?`${code} - ${name}`:"Chưa xác định";
      }
      return {...row, _sector:label, _xa:String(row[targetXa]||"Khác").trim()};
    });
    let finalRows:any[]=[];
    if(reportType==="pivot"){
      const xas=[...new Set(processed.map(r=>r._xa))].sort();
      const sectors=[...new Set(processed.map(r=>r._sector))].sort();
      xas.forEach(xa=>{
        const obj:any={"Địa bàn":xa};
        sectors.forEach(sec=>{
          const match=processed.filter(r=>r._xa===xa && r._sector===sec);
          let dt=0,ld=0;
          match.forEach(r=>{
            if(qrDoanhthu) dt+=Number(r[qrDoanhthu])||0;
            if(qrLaodong) ld+=Number(r[qrLaodong])||0;
          });
          obj[`${sec} - DT`]=Math.round(dt*100)/100;
          obj[`${sec} - LĐ`]=Math.round(ld);
        });
        finalRows.push(obj);
      });
    }else{
      const groups=new Map();
      processed.forEach(r=>{const key=JSON.stringify({s:r._sector,x:r._xa}); if(!groups.has(key)) groups.set(key,[]); groups.get(key).push(r);});
      groups.forEach((arr,key)=>{
        const {s,x}=JSON.parse(key);
        let dt=0,ld=0;
        arr.forEach(r=>{dt+=Number(r[qrDoanhthu])||0; ld+=Number(r[qrLaodong])||0;});
        finalRows.push({[`Ngành cấp ${level}`]:s, "Địa bàn":x, "Số DN":arr.length, "Doanh thu":Math.round(dt*100)/100, "Lao động":Math.round(ld)});
      });
    }
    setQuickRows(finalRows); setQuickCols(Object.keys(finalRows[0]||{})); setQuickLevel(level);
    setLoading(false); setMsg(`Báo cáo cấp ${level} hoàn tất`);
  };

  // Chuẩn hóa mã ngành VSIC
  const handleStandardize = async()=>{
    if(!mainData.length||!stdCol){alert("Chọn cột mã ngành"); return;}
    setLoading(true);
    let valid=0,invalid=0;
    const updated=await chunkProcess(mainData,5000,row=>{
      const raw=row[stdCol];
      const clean=normalizeSectorCode(raw);
      const lookup=lookupSectorNameWithFallback(clean);
      if(lookup.exactMatched) valid++; else invalid++;
      const newRow:any={};
      Object.keys(row).forEach(k=>{
        newRow[k]=row[k];
        if(k===stdCol){ newRow["Tên ngành chuẩn"]=lookup.name||"(Không có)"; newRow["Trạng thái"]=lookup.exactMatched?"✅ Hợp lệ":"❌ Lỗi"; }
      });
      if(!newRow["Tên ngành chuẩn"]){ newRow["Tên ngành chuẩn"]=lookup.name; newRow["Trạng thái"]=lookup.exactMatched?"✅ Hợp lệ":"❌ Lỗi"; }
      return newRow;
    });
    setMainData(updated); setColumns(Object.keys(updated[0]||{}));
    setStdMatch({total:updated.length,valid,invalid});
    setLoading(false); setMsg(`Chuẩn hóa xong: ${valid} hợp lệ, ${invalid} lỗi`);
  };

  // Đối chiếu 2 cột bất kỳ
  const handleCrossCompare = async()=>{
    if(!mainData.length||!compA||!compB){alert("Chọn 2 cột để so sánh"); return;}
    setLoading(true);
    let match=0,mismatch=0;
    const updated=await chunkProcess(mainData,5000,row=>{
      let a=String(row[compA]??"").trim(); let b=String(row[compB]??"").trim();
      let ok=false;
      if(compRule==="exact") ok=a===b;
      else if(compRule==="normalize") ok=a.toLowerCase().replace(/\s+/g," ")===b.toLowerCase().replace(/\s+/g," ");
      else if(compRule==="sector_code"){
        const ca=a.replace(/\D/g,""); const cb=b.replace(/\D/g,"");
        ok=(ca===cb&&ca!=="")||(ca&&cb&&(ca.startsWith(cb)||cb.startsWith(ca)));
      }else ok=a.toLowerCase().includes(b.toLowerCase())||b.toLowerCase().includes(a.toLowerCase());
      if(ok) match++; else mismatch++;
      const newRow:any={};
      Object.keys(row).forEach(k=>{
        newRow[k]=row[k];
        if(k===compB) newRow["Kết quả đối chiếu"]=ok?"✅ Khớp":"❌ Lệch";
      });
      if(!newRow["Kết quả đối chiếu"]) newRow["Kết quả đối chiếu"]=ok?"✅ Khớp":"❌ Lệch";
      return newRow;
    });
    setMainData(updated); setColumns(Object.keys(updated[0]||{}));
    setLoading(false); setMsg(`Đối chiếu: ${match} khớp, ${mismatch} lệch`);
  };

  // Thêm cột cấp 1 và cấp 2 vào dữ liệu chính
  const handleAppendLevels = async()=>{
    if(!mainData.length){alert("Chưa có dữ liệu"); return;}
    const target=qrManganh||mapping.manganh;
    if(!target){alert("Chọn cột mã ngành"); return;}
    setLoading(true);
    const updated=await chunkProcess(mainData,5000,row=>{
      const raw=row[target];
      const mng=normalizeSectorCode(raw);
      let c1="", n1="", c2="", n2="";
      if(mng){
        c2=mng.slice(0,2); n2=userSectorMap.get(c2)||"";
        if(/^[A-Z]$/.test(mng)) c1=mng.toUpperCase();
        else c1=getParentSectorCode(mng)||"";
        n1=userSectorMap.get(c1)||"";
      }
      const newRow:any={};
      Object.keys(row).forEach(k=>{
        newRow[k]=row[k];
        if(k===target){
          newRow["Mã cấp 1"]=c1; newRow["Tên cấp 1"]=n1;
          newRow["Mã cấp 2"]=c2; newRow["Tên cấp 2"]=n2;
        }
      });
      if(!newRow["Mã cấp 1"]){ newRow["Mã cấp 1"]=c1; newRow["Tên cấp 1"]=n1; newRow["Mã cấp 2"]=c2; newRow["Tên cấp 2"]=n2; }
      return newRow;
    });
    setMainData(updated); setColumns(Object.keys(updated[0]||{}));
    setLoading(false); setMsg("Đã thêm cột cấp 1 và cấp 2");
  };

  // Tính toán cột mới
  const handleCalcColumn = async()=>{
    if(!calcName.trim()){alert("Nhập tên cột mới"); return;}
    if(!calcCol1){alert("Chọn cột thứ nhất"); return;}
    if(calcType==="column"&&!calcCol2){alert("Chọn cột thứ hai"); return;}
    if(calcType==="constant"&&!calcConst.trim()){alert("Nhập hằng số"); return;}
    setLoading(true);
    const compute=(a:any,b:any)=>{
      if(calcOp==="concat") return `${a??""} ${b??""}`.trim();
      const na=Number(String(a).replace(/[^0-9.-]/g,""))||0;
      const nb=typeof b==="number"?b:(Number(String(b).replace(/[^0-9.-]/g,""))||0);
      let res=0;
      if(calcOp==="+") res=na+nb; else if(calcOp==="-") res=na-nb;
      else if(calcOp==="*") res=na*nb; else if(calcOp==="/") res=nb!==0?na/nb:0;
      if(calcRound==="int") return Math.round(res);
      if(calcRound==="1dec") return Math.round(res*10)/10;
      if(calcRound==="2dec") return Math.round(res*100)/100;
      return res;
    };
    const updatedRaw=await chunkProcess(rawData,5000,row=>({...row,[calcName]:compute(row[calcCol1], calcType==="column"?row[calcCol2]:calcConst)}));
    const updatedMain=await chunkProcess(mainData,5000,row=>({...row,[calcName]:compute(row[calcCol1], calcType==="column"?row[calcCol2]:calcConst)}));
    const newCols=columns.includes(calcName)?columns:[...columns,calcName];
    let newConfigs=[...colConfigs];
    if(!newConfigs.some(c=>c.originalName===calcName)) newConfigs.push({originalName:calcName, use:true, newName:calcName, role:""});
    setRawData(updatedRaw); setMainData(updatedMain); setColumns(newCols); setColConfigs(newConfigs);
    setCalcName(""); setCalcCol1(""); setCalcCol2(""); setCalcConst("");
    setLoading(false); setMsg(`Đã thêm cột "${calcName}"`);
  };

  // Xuất Excel kết quả
  const handleExport = ()=>{
    const ws=XLSX.utils.json_to_sheet(mainData);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"KetQua");
    XLSX.writeFile(wb,fileName||"KetQua.xlsx");
  };

  // ==================== RESTORE SESSION & AUTO SAVE ====================
  useEffect(()=>{
    const savedCat=localStorage.getItem("custom_vsic_data");
    if(savedCat){
      const obj=JSON.parse(savedCat);
      const m=new Map(Object.entries(obj));
      setUserSectorMap(m);
      clearAllSectorsInVSIC();
      loadSectorsIntoVSIC(obj);
      setUserSectorFile(localStorage.getItem("custom_vsic_filename")||"");
    }
  },[]);

  // ==================== RENDER ====================
  // Filter dữ liệu cho tab xem
  const filteredMain = useMemo(()=>{
    if(!searchTerm) return mainData;
    const t=searchTerm.toLowerCase();
    return mainData.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(t)));
  },[mainData,searchTerm]);
  const totalPagesView = Math.ceil(filteredMain.length/pageSize);
  const paginatedMain = filteredMain.slice((pageView-1)*pageSize, pageView*pageSize);

  // Auth gate
  if(!isAuthorized){
    return (
      <div className="min-h-screen bg-gradient-to-br from-[#0a0f1f] to-[#111827] flex items-center justify-center">
        <div className="bg-[#1f2937]/90 p-8 rounded-2xl w-96 border border-purple-500/30">
          <div className="text-center"><Lock className="w-12 h-12 text-purple-500 mx-auto"/><h2 className="text-2xl font-bold mt-2">BẢO MẬT</h2></div>
          <form onSubmit={(e)=>{e.preventDefault(); if(typedPwd===appPwd){localStorage.setItem("vsic_app_authorized","true"); setIsAuthorized(true);}else setPwdErr("Sai mật khẩu");}}>
            <input type="password" value={typedPwd} onChange={e=>setTypedPwd(e.target.value)} className="w-full bg-[#111827] border border-gray-700 rounded-lg p-2 mt-4" placeholder="Mật khẩu" autoFocus/>
            {pwdErr && <p className="text-red-400 text-sm mt-1">{pwdErr}</p>}
            <button type="submit" className="w-full mt-4 bg-purple-600 py-2 rounded-lg">Xác nhận</button>
          </form>
          <p className="text-center text-xs text-gray-500 mt-4">Mật khẩu: admin123</p>
        </div>
      </div>
    );
  }

  // Main UI
  return (
    <div className="flex flex-col h-screen bg-gradient-to-br from-[#0a0f1f] to-[#111827] text-white overflow-hidden">
      <header className="bg-[#1f2937]/80 backdrop-blur border-b border-gray-800 px-6 py-3 flex justify-between items-center">
        <div className="flex items-center gap-2"><Layers className="w-6 h-6 text-purple-500"/><h1 className="text-xl font-bold">HỆ THỐNG VSIC</h1></div>
        <div className="flex gap-3">
          <button onClick={()=>{localStorage.removeItem("vsic_app_authorized"); setIsAuthorized(false);}} className="bg-red-950/40 px-3 py-1 rounded-lg text-sm"><LogOut className="w-4 h-4 inline"/> Thoát</button>
          {fileName && <div className="bg-[#111827] px-3 py-1 rounded-lg text-sm">{fileName} ({mainData.length})</div>}
        </div>
      </header>
      <div className="flex flex-1 overflow-hidden">
        <aside className="w-64 bg-[#1f2937]/50 border-r border-gray-800 p-4 space-y-1 overflow-y-auto">
          {[
            ["trangchu","🏠","Trang chủ"],[ "xemdulieu","📂","Xem & Định nghĩa cột"],[ "ghepnoi","🌿","Ghép nối"],[ "sosanh","🔍","So sánh"],[ "tachfile","✂️","Tách file"],[ "tonghop","📊","Tổng hợp báo cáo"],[ "bieudotrucquan","📈","Biểu đồ"],[ "chuanhoanganh","🧠","Chuẩn hóa VSIC"],[ "doichieumota","🔄","Đối chiếu mô tả"],[ "danhmucvsic","📚","Danh mục ngành"]
          ].map(([id,icon,label])=>(
            <button key={id} onClick={()=>setActiveTab(id)} className={`w-full flex items-center gap-2 px-3 py-2 rounded-xl text-sm ${activeTab===id?"bg-purple-600/30 text-purple-400":"hover:bg-white/5"}`}>{icon} {label}</button>
          ))}
        </aside>
        <main className="flex-1 overflow-y-auto p-6">
          {loading && (
            <div className="fixed inset-0 bg-black/70 flex items-center justify-center z-50"><div className="bg-[#1f2937] p-6 rounded-2xl w-80 text-center"><Loader2 className="animate-spin w-10 h-10 mx-auto"/><p>{msg}</p><div className="w-full bg-gray-700 rounded-full h-1 mt-3"><div className="bg-purple-600 h-1 rounded-full" style={{width:`${progress}%`}}></div></div></div></div>
          )}
          {/* TRANG CHỦ */}
          {activeTab==="trangchu" && (
            <div><div className="bg-gradient-to-r from-purple-900/30 to-[#1f2937] p-6 rounded-2xl"><h2 className="text-2xl font-bold">Hệ thống phân tích dữ liệu ngành</h2><p className="text-gray-400">Upload danh mục ngành, chuẩn hóa, ghép nối, báo cáo</p></div><div className="grid md:grid-cols-2 gap-4 mt-6"><div className="bg-[#1f2937]/50 p-4 rounded-xl">📁 Bước 1: Tải dữ liệu chính</div><div className="bg-[#1f2937]/50 p-4 rounded-xl">📚 Bước 2: Nạp danh mục ngành</div><div className="bg-[#1f2937]/50 p-4 rounded-xl">⚙️ Bước 3: Xử lý & báo cáo</div><div className="bg-[#1f2937]/50 p-4 rounded-xl">📊 Kết quả hiển thị ngay tại tab</div></div></div>
          )}
          {/* XEM & ĐỊNH NGHĨA CỘT */}
          {activeTab==="xemdulieu" && (
            <div className="space-y-6">
              <div className="bg-[#1f2937] rounded-2xl p-6"><div className="flex justify-between"><h3 className="text-lg font-bold">📂 FILE DỮ LIỆU CHÍNH</h3><label className="bg-purple-600 px-4 py-2 rounded-xl cursor-pointer"><FileUp className="w-4 h-4 inline"/> TẢI FILE<input type="file" accept=".xlsx,.xls,.csv" onChange={(e)=>handleFileUpload(e,"main")} className="hidden"/></label></div>{detectedSheets.length>1 && (<div className="mt-3 p-3 bg-amber-950/30 rounded-lg"><p>Ghép {detectedSheets.length} sheets</p><div className="flex gap-2">{detectedSheets.map(s=><label key={s}><input type="checkbox" checked={selectedSheets.includes(s)} onChange={()=>setSelectedSheets(prev=>prev.includes(s)?prev.filter(x=>x!==s):[...prev,s])}/> {s}</label>)}</div><select value={sheetKey} onChange={e=>setSheetKey(e.target.value)} className="bg-[#111827] p-1 rounded"><option value="">Chọn cột khóa</option>{columns.map(c=><option key={c}>{c}</option>)}</select><button onClick={handleMergeSheets} className="bg-amber-600 px-3 py-1 rounded mt-2">Ghép sheets</button></div>)}</div>
              {rawData.length>0 && (
                <div className="bg-[#1f2937] rounded-2xl p-6"><div className="flex justify-between"><h4>📌 Định nghĩa lại cột</h4><button onClick={()=>setExpanded(!expanded)}>{expanded?"Thu gọn":"Mở rộng"}</button></div>{expanded && (<div className="overflow-x-auto"><table className="w-full text-xs"><thead><tr><th>Dùng</th><th>Tên gốc</th><th>Tên mới</th><th>Vai trò</th></tr></thead><tbody>{colConfigs.map((cfg,i)=><tr key={i}><td><input type="checkbox" checked={cfg.use} onChange={e=>{const newC=[...colConfigs]; newC[i].use=e.target.checked; setColConfigs(newC);}}/></td><td>{cfg.originalName}</td><td><input value={cfg.newName} onChange={e=>{const newC=[...colConfigs]; newC[i].newName=e.target.value; setColConfigs(newC);}} className="bg-[#111827] border border-gray-700 rounded px-1"/></td><td><select value={cfg.role} onChange={e=>{const newC=[...colConfigs]; newC[i].role=e.target.value as any; setColConfigs(newC);}}><option value="">--</option><option value="mota">Mô tả</option><option value="manganh">Mã ngành</option><option value="xa">Xã</option><option value="doanhthu">Doanh thu</option><option value="laodong">Lao động</option></select></td></tr>)}</tbody></table><button onClick={handleRedefine} className="mt-3 bg-indigo-600 px-4 py-1 rounded">Áp dụng</button></div>)}</div>
              )}
              {rawData.length>0 && (
                <div className="bg-[#1f2937] rounded-2xl p-6"><h4 className="text-indigo-400 font-bold">🧮 TÍNH TOÁN CỘT MỚI</h4><div className="grid grid-cols-2 md:grid-cols-12 gap-2 mt-2"><input className="col-span-3 bg-[#111827] border rounded p-1" placeholder="Tên cột mới" value={calcName} onChange={e=>setCalcName(e.target.value)}/><select className="col-span-2 bg-[#111827] border rounded p-1" value={calcCol1} onChange={e=>setCalcCol1(e.target.value)}><option value="">Cột A</option>{columns.map(c=><option key={c}>{c}</option>)}</select><select className="col-span-1 bg-[#111827] border rounded p-1" value={calcOp} onChange={e=>setCalcOp(e.target.value as any)}><option value="+">+</option><option value="-">-</option><option value="*">*</option><option value="/">/</option><option value="concat">concat</option></select><select className="col-span-1 bg-[#111827] border rounded p-1" value={calcType} onChange={e=>setCalcType(e.target.value as any)}><option value="column">Cột</option><option value="constant">Số</option></select>{calcType==="column"?<select className="col-span-3 bg-[#111827] border rounded p-1" value={calcCol2} onChange={e=>setCalcCol2(e.target.value)}><option value="">Cột B</option>{columns.map(c=><option key={c}>{c}</option>)}</select>:<input className="col-span-3 bg-[#111827] border rounded p-1" placeholder="Hằng số" value={calcConst} onChange={e=>setCalcConst(e.target.value)}/>}</div><div className="flex gap-3 mt-2"><select value={calcRound} onChange={e=>setCalcRound(e.target.value as any)} className="bg-[#111827] border rounded p-1"><option value="none">Không làm tròn</option><option value="int">Số nguyên</option><option value="1dec">1 số thập phân</option><option value="2dec">2 số thập phân</option></select><button onClick={handleCalcColumn} className="bg-indigo-600 px-4 py-1 rounded">➕ Thêm cột</button></div></div>
              )}
              <DataPreviewTable data={mainData} columns={columns} title="DỮ LIỆU CHÍNH" />
            </div>
          )}
          {/* GHÉP NỐI */}
          {activeTab==="ghepnoi" && (
            <div className="space-y-6"><div className="bg-[#1f2937] rounded-2xl p-6"><div className="grid md:grid-cols-2 gap-4"><div><h4>Bảng trái</h4><label className="bg-gray-800 px-3 py-1 rounded cursor-pointer">Chọn file<input type="file" onChange={(e)=>handleFileUpload(e,"left")} className="hidden"/></label><div>{leftName}</div>{leftData.length>0 && <select value={leftKey} onChange={e=>setLeftKey(e.target.value)} className="w-full mt-2 bg-[#111827] border rounded p-1"><option value="">Khóa trái</option>{Object.keys(leftData[0]||{}).map(c=><option key={c}>{c}</option>)}</select>}</div><div><h4>Bảng phải</h4><label className="bg-gray-800 px-3 py-1 rounded cursor-pointer">Chọn file<input type="file" onChange={(e)=>handleFileUpload(e,"right")} className="hidden"/></label><div>{rightName}</div>{rightData.length>0 && <select value={rightKey} onChange={e=>setRightKey(e.target.value)} className="w-full mt-2 bg-[#111827] border rounded p-1"><option value="">Khóa phải</option>{Object.keys(rightData[0]||{}).map(c=><option key={c}>{c}</option>)}</select>}</div></div><button onClick={handleMerge} className="bg-blue-600 px-5 py-2 rounded-xl mt-4">Ghép nối</button></div><DataPreviewTable data={mainData} columns={columns} title="KẾT QUẢ GHÉP NỐI"/></div>
          )}
          {/* SO SÁNH */}
          {activeTab==="sosanh" && (
            <div className="space-y-6"><div className="bg-[#1f2937] rounded-2xl p-6"><div className="grid md:grid-cols-2 gap-4"><div><h4>File cũ</h4><label className="bg-gray-800 px-3 py-1 rounded cursor-pointer">Chọn file<input type="file" onChange={(e)=>handleFileUpload(e,"old")} className="hidden"/></label><div>{oldName}</div></div><div><h4>File mới</h4><label className="bg-gray-800 px-3 py-1 rounded cursor-pointer">Chọn file<input type="file" onChange={(e)=>handleFileUpload(e,"new")} className="hidden"/></label><div>{newName}</div></div></div>{oldData.length&&newData.length&&<div className="mt-3"><select value={diffKey} onChange={e=>setDiffKey(e.target.value)} className="bg-[#111827] border rounded p-1"><option value="">Chọn cột khóa chung</option>{Object.keys(oldData[0]||{}).filter(c=>Object.keys(newData[0]||{}).includes(c)).map(c=><option key={c}>{c}</option>)}</select></div>}<button onClick={handleCompare} className="bg-cyan-600 px-5 py-2 rounded-xl mt-4">So sánh</button></div><DataPreviewTable data={mainData} columns={columns} title="KẾT QUẢ SO SÁNH"/></div>
          )}
          {/* TÁCH FILE */}
          {activeTab==="tachfile" && (
            <div className="space-y-6"><div className="bg-[#1f2937] rounded-2xl p-6"><h3>✂️ Tách file theo cột</h3><select value={splitCol} onChange={e=>setSplitCol(e.target.value)} className="bg-[#111827] border rounded p-1 mt-2"><option value="">Chọn cột</option>{columns.map(c=><option key={c}>{c}</option>)}</select><button onClick={handleSplit} className="bg-pink-600 px-5 py-2 rounded-xl mt-4 ml-3">Tách & tải ZIP</button></div></div>
          )}
          {/* TỔNG HỢP BÁO CÁO */}
          {activeTab==="tonghop" && (
            <div className="space-y-6"><div className="bg-[#1f2937] rounded-2xl p-6"><h3>📊 Báo cáo theo ngành & xã</h3><div className="grid md:grid-cols-2 gap-3 mt-3"><select value={qrManganh} onChange={e=>setQrManganh(e.target.value)} className="bg-[#111827] border rounded p-1"><option value="">Cột mã ngành</option>{columns.map(c=><option key={c}>{c}</option>)}</select><select value={qrXa} onChange={e=>setQrXa(e.target.value)} className="bg-[#111827] border rounded p-1"><option value="">Cột xã</option>{columns.map(c=><option key={c}>{c}</option>)}</select><select value={qrDoanhthu} onChange={e=>setQrDoanhthu(e.target.value)} className="bg-[#111827] border rounded p-1"><option value="">Doanh thu (tùy chọn)</option>{columns.map(c=><option key={c}>{c}</option>)}</select><select value={qrLaodong} onChange={e=>setQrLaodong(e.target.value)} className="bg-[#111827] border rounded p-1"><option value="">Lao động (tùy chọn)</option>{columns.map(c=><option key={c}>{c}</option>)}</select></div><div className="flex gap-3 mt-3"><button onClick={()=>handleQuickReport(1)} className="bg-emerald-600 px-4 py-1 rounded">Báo cáo cấp 1</button><button onClick={()=>handleQuickReport(2)} className="bg-emerald-600 px-4 py-1 rounded">Báo cáo cấp 2</button><button onClick={handleAppendLevels} className="bg-blue-600 px-4 py-1 rounded">Thêm cấp 1,2</button></div></div>{quickRows.length>0 && <BeautifulReportTable rows={quickRows} cols={quickCols} level={quickLevel} reportType={reportType} onExport={()=>{const ws=XLSX.utils.json_to_sheet(quickRows); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"BaoCao"); XLSX.writeFile(wb,"BaoCao.xlsx");}} />}<DataPreviewTable data={mainData} columns={columns} title="DỮ LIỆU HIỆN TẠI"/></div>
          )}
          {/* BIỂU ĐỒ */}
          {activeTab==="bieudotrucquan" && <SectorRevenueChart mainData={mainData} columns={columns} mapping={mapping} />}
          {/* CHUẨN HÓA VSIC */}
          {activeTab==="chuanhoanganh" && (
            <div className="space-y-6"><div className="bg-[#1f2937] rounded-2xl p-6"><h3>🧠 Chuẩn hóa mã ngành</h3><select value={stdCol} onChange={e=>setStdCol(e.target.value)} className="bg-[#111827] border rounded p-1 mt-2"><option value="">Chọn cột mã ngành</option>{columns.map(c=><option key={c}>{c}</option>)}</select><button onClick={handleStandardize} className="ml-3 bg-indigo-600 px-4 py-1 rounded">Chuẩn hóa</button>{stdMatch.total>0 && <div className="mt-2 text-sm">✅ Hợp lệ: {stdMatch.valid} | ❌ Lỗi: {stdMatch.invalid}</div>}<div className="mt-3"><h4 className="font-bold">Đối chiếu 2 cột</h4><select value={compA} onChange={e=>setCompA(e.target.value)} className="bg-[#111827] border rounded p-1"><option value="">Cột A</option>{columns.map(c=><option key={c}>{c}</option>)}</select><select value={compB} onChange={e=>setCompB(e.target.value)} className="ml-2 bg-[#111827] border rounded p-1"><option value="">Cột B</option>{columns.map(c=><option key={c}>{c}</option>)}</select><select value={compRule} onChange={e=>setCompRule(e.target.value)} className="ml-2 bg-[#111827] border rounded p-1"><option value="exact">Chính xác</option><option value="normalize">Chuẩn hóa</option><option value="sector_code">Mã ngành</option><option value="substring">Chứa chuỗi</option></select><button onClick={handleCrossCompare} className="ml-2 bg-cyan-600 px-3 py-1 rounded">Đối chiếu</button></div></div><DataPreviewTable data={mainData} columns={columns} title="SAU CHUẨN HÓA"/></div>
          )}
          {/* ĐỐI CHIẾU MÔ TẢ */}
          {activeTab==="doichieumota" && <DescriptorMatchScanner mainData={mainData} columns={columns} mapping={mapping} />}
          {/* DANH MỤC NGÀNH */}
          {activeTab==="danhmucvsic" && (
            <div className="space-y-6"><div className="bg-[#1f2937] rounded-2xl p-6"><div className="flex justify-between"><div><h3>📚 Nạp danh mục ngành của bạn</h3><p className="text-sm text-gray-400">Excel/CSV (mã, tên)</p></div><div><label className="bg-cyan-600 px-4 py-2 rounded-xl cursor-pointer"><FileUp className="w-4 h-4 inline"/> TẢI FILE<input type="file" accept=".xlsx,.xls,.csv" onChange={handleUploadCatalog} className="hidden"/></label><button onClick={handleClearCatalog} className="ml-2 bg-red-800/60 px-4 py-2 rounded-xl">XÓA</button></div></div>{userSectorFile && <div className="mt-3 bg-green-900/30 p-2 rounded">✅ Đã nạp: {userSectorFile} ({userSectorMap.size} mã)</div>}</div><VsicCatalogExplorer /></div>
          )}
        </main>
      </div>
    </div>
  );
}
    
