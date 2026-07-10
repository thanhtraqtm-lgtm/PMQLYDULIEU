import React, { useState, useEffect, useRef } from "react";
import { Key, Shield, ShieldCheck, RefreshCw, Smartphone, Check, Lock, Cpu, Eye, EyeOff } from "lucide-react";

interface SignatureTokenProps {
  onSave: (dataUrl: string, hash: string, certInfo: string) => void;
  onClear: () => void;
  mst: string;
  enterpriseName: string;
  representative: string;
}

const CA_PROVIDERS = [
  { id: "vnpt", name: "VNPT-CA (Tập đoàn Bưu chính Viễn thông Việt Nam)", brand: "VNPT" },
  { id: "viettel", name: "Viettel-CA (Tập đoàn Công nghiệp - Viễn thông Quân đội)", brand: "Viettel" },
  { id: "fpt", name: "FPT-CA (Công ty Cổ phần Hệ thống Thông tin FPT)", brand: "FPT" },
  { id: "bkav", name: "Bkav-CA (Tập đoàn Công nghệ Bkav)", brand: "Bkav" },
  { id: "misa", name: "MISA eSign (Công ty Cổ phần MISA)", brand: "MISA" },
];

export const SignatureToken: React.FC<SignatureTokenProps> = ({
  onSave,
  onClear,
  mst = "0900123456",
  enterpriseName = "CÔNG TY TNHH THƯƠNG MẠI HƯNG YÊN",
  representative = "Nguyễn Văn A"
}) => {
  const [provider, setProvider] = useState("vnpt");
  const [isScanning, setIsScanning] = useState(false);
  const [scanned, setScanned] = useState(false);
  const [pin, setPin] = useState("12345678");
  const [showPin, setShowPin] = useState(false);
  const [isSigning, setIsSigning] = useState(false);
  const [signed, setSigned] = useState(false);
  const [signedTime, setSignedTime] = useState("");
  const [shaHash, setShaHash] = useState("");
  const [serialNumber, setSerialNumber] = useState("");
  const [errorMsg, setErrorMsg] = useState("");

  const canvasRef = useRef<HTMLCanvasElement | null>(null);

  // Sinh ngẫu nhiên số Serial khi quét thiết bị
  useEffect(() => {
    if (scanned && !serialNumber) {
      const hex = "0123456789ABCDEF";
      let res = "5401A";
      for (let i = 0; i < 15; i++) {
        res += hex[Math.floor(Math.random() * 16)];
      }
      setSerialNumber(res);
    }
  }, [scanned, serialNumber]);

  const handleScanToken = () => {
    setIsScanning(true);
    setErrorMsg("");
    setTimeout(() => {
      setIsScanning(false);
      setScanned(true);
    }, 1500); // Giả lập quét USB 1.5s
  };

  const generateRedSealImage = (
    entName: string,
    taxCode: string,
    rep: string,
    caName: string,
    serial: string,
    dateStr: string
  ): string => {
    // Tạo canvas tạm để vẽ dấu mộc đỏ điện tử của Việt Nam
    const canvas = document.createElement("canvas");
    canvas.width = 450;
    canvas.height = 150;
    const ctx = canvas.getContext("2d");
    if (!ctx) return "";

    // Vẽ nền trắng
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // Vẽ khung viền kép đỏ chữ nhật nét mảnh (đặc trưng chữ ký số VN)
    ctx.strokeStyle = "#e11d48"; // màu đỏ rose-600 cực đẹp
    ctx.lineWidth = 3;
    ctx.strokeRect(6, 6, canvas.width - 12, canvas.height - 12);
    
    ctx.lineWidth = 1;
    ctx.strokeRect(12, 12, canvas.width - 24, canvas.height - 24);

    // Vẽ text
    ctx.fillStyle = "#e11d48";
    
    // Tiêu đề chữ ký số
    ctx.font = "bold 13px Arial, Helvetica, sans-serif";
    ctx.fillText("Signature Verified", 24, 35);
    
    // Tên doanh nghiệp ký
    ctx.font = "bold 12px Arial, Helvetica, sans-serif";
    const displayName = entName.length > 40 ? entName.substring(0, 38) + "..." : entName;
    ctx.fillText(`Ký bởi: ${displayName.toUpperCase()}`, 24, 58);
    
    // Mã số thuế
    ctx.font = "11px Arial, Helvetica, sans-serif";
    ctx.fillText(`Mã số doanh nghiệp: ${taxCode || "0900123456"}`, 24, 76);
    
    // Người đại diện pháp lý
    ctx.fillText(`Người đại diện: ${rep || "Người đại diện"}`, 24, 94);
    
    // Ngày giờ ký và Chứng thư số phát hành bởi ai
    ctx.font = "italic 10px Arial, Helvetica, sans-serif";
    ctx.fillText(`Thời gian ký: ${dateStr}`, 24, 114);
    
    const caBrand = CA_PROVIDERS.find(c => c.id === provider)?.brand || "VNPT-CA";
    ctx.fillText(`Chứng thư số bảo mật cấp bởi: ${caBrand} (Serial: ${serial})`, 24, 131);

    return canvas.toDataURL("image/png");
  };

  const handleSign = () => {
    if (!pin) {
      setErrorMsg("Vui lòng nhập mã PIN bảo mật của USB Token!");
      return;
    }
    if (pin.length < 4) {
      setErrorMsg("Mã PIN không hợp lệ (tối thiểu 4 ký tự)!");
      return;
    }

    setIsSigning(true);
    setErrorMsg("");

    setTimeout(() => {
      const now = new Date();
      const dateStr = now.toLocaleString("vi-VN", {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        day: "2-digit",
        month: "2-digit",
        year: "numeric"
      });

      // Sinh mã băm bảo mật SHA-256 ngẫu nhiên đại diện cho chữ ký số
      const hex = "0123456789ABCDEF";
      let signHash = "SHA256-DIGITAL-";
      for (let i = 0; i < 24; i++) {
        signHash += hex[Math.floor(Math.random() * 16)];
      }

      const activeCa = CA_PROVIDERS.find(c => c.id === provider)?.name || "VNPT-CA";
      const certInfoText = `Ký số bởi: ${enterpriseName} | MST: ${mst} | Đại diện: ${representative} | CA: ${activeCa} | Serial: ${serialNumber} | Ngày ký: ${dateStr}`;

      // Vẽ mộc đỏ
      const redSealBase64 = generateRedSealImage(
        enterpriseName || "DOANH NGHIỆP KHẢO SÁT",
        mst || "0900112233",
        representative || "Người đại diện",
        activeCa,
        serialNumber,
        dateStr
      );

      setIsSigning(false);
      setSigned(true);
      setSignedTime(dateStr);
      setShaHash(signHash);

      // Lưu lại
      onSave(redSealBase64, signHash, certInfoText);
    }, 1200);
  };

  const handleReset = () => {
    setScanned(false);
    setSigned(false);
    setSignedTime("");
    setShaHash("");
    setSerialNumber("");
    setErrorMsg("");
    onClear();
  };

  return (
    <div className="border border-slate-300 rounded-lg overflow-hidden bg-white shadow-xs font-sans">
      {/* Header */}
      <div className="bg-slate-900 text-white p-3 flex items-center justify-between text-xs font-bold uppercase tracking-wider">
        <div className="flex items-center gap-1.5">
          <Key className="w-4 h-4 text-amber-400 animate-pulse" />
          <span>Mô phỏng Chữ ký số USB Token Doanh nghiệp</span>
        </div>
        <span className="text-[9px] px-2 py-0.5 bg-amber-500 text-slate-950 font-mono rounded-full font-black">
          BẢO MẬT SSL
        </span>
      </div>

      <div className="p-4 space-y-4 text-xs">
        {/* Bước 1: Chọn Nhà cung cấp CA */}
        {!scanned && (
          <div className="space-y-2">
            <label className="block font-bold text-slate-700">1. Chọn nhà cung cấp Chứng thư số (CA):</label>
            <select
              value={provider}
              onChange={(e) => setProvider(e.target.value)}
              className="w-full bg-slate-50 border border-slate-300 rounded-md p-2 text-xs focus:ring-1 focus:ring-indigo-500 focus:outline-none font-medium text-slate-800"
            >
              {CA_PROVIDERS.map((ca) => (
                <option key={ca.id} value={ca.id}>
                  {ca.name}
                </option>
              ))}
            </select>

            {/* Nút quét token */}
            <div className="pt-2">
              <button
                type="button"
                onClick={handleScanToken}
                disabled={isScanning}
                className="w-full bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-300 text-white font-bold py-2 px-4 rounded-md shadow-xs transition-all active:scale-98 flex items-center justify-center gap-2 cursor-pointer"
              >
                {isScanning ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin text-white" />
                    Đang quét thiết bị USB Token trên máy tính...
                  </>
                ) : (
                  <>
                    <Cpu className="w-4 h-4 text-amber-300" />
                    Quét kết nối thiết bị USB Token CA
                  </>
                )}
              </button>
              <p className="text-[10px] text-slate-500 italic mt-1.5 text-center">
                Vui lòng đảm bảo USB Token của {CA_PROVIDERS.find(c => c.id === provider)?.brand} đã được cắm vào máy tính của bạn.
              </p>
            </div>
          </div>
        )}

        {/* Bước 2: Hiển thị chứng thư số đã quét & Nhập mã PIN để Ký */}
        {scanned && !signed && (
          <div className="space-y-3 animate-fadeIn">
            {/* Box thông tin CTS */}
            <div className="border border-amber-200 bg-amber-50/45 rounded-lg p-3 space-y-2">
              <div className="flex items-center gap-1 text-[11px] font-bold text-amber-900 border-b border-amber-200 pb-1">
                <Shield className="w-3.5 h-3.5 text-amber-600" />
                PHÁT HIỆN THIẾT BỊ CHỨNG THƯ SỐ DOANH NGHIỆP:
              </div>
              <div className="space-y-1 text-slate-800 text-[11px]">
                <p>
                  <strong className="text-slate-900 font-serif">Doanh nghiệp:</strong> {enterpriseName.toUpperCase()}
                </p>
                <p>
                  <strong className="text-slate-900">Mã số thuế:</strong> {mst || "0900123456"}
                </p>
                <p>
                  <strong className="text-slate-900">Số Seri thiết bị:</strong> <span className="font-mono font-bold text-indigo-700 bg-slate-100 px-1 py-0.5 rounded-xs">{serialNumber}</span>
                </p>
                <p>
                  <strong className="text-slate-900">Nhà phát hành CA:</strong> {CA_PROVIDERS.find(c => c.id === provider)?.name}
                </p>
                <p>
                  <strong className="text-slate-900">Người đại diện pháp lý:</strong> {representative || "Chưa xác định"}
                </p>
              </div>
            </div>

            {/* Nhập mã PIN */}
            <div className="space-y-2">
              <div className="flex justify-between items-center">
                <label className="block font-bold text-slate-700 flex items-center gap-1">
                  <Lock className="w-3 h-3 text-indigo-600" />
                  Nhập mã PIN USB Token để ký báo cáo:
                </label>
                <span className="text-[9px] text-slate-500 italic">Mặc định: 12345678</span>
              </div>
              <div className="relative">
                <input
                  type={showPin ? "text" : "password"}
                  value={pin}
                  onChange={(e) => setPin(e.target.value)}
                  placeholder="Nhập mã PIN chữ ký số..."
                  className="w-full bg-white border border-slate-300 rounded-md py-1.5 pl-3 pr-10 text-xs font-mono font-bold focus:ring-1 focus:ring-emerald-500 focus:outline-none"
                />
                <button
                  type="button"
                  onClick={() => setShowPin(!showPin)}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer"
                >
                  {showPin ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>
              {errorMsg && (
                <p className="text-red-600 text-[10px] font-bold bg-red-50 p-1.5 rounded border border-red-100">
                  {errorMsg}
                </p>
              )}
            </div>

            {/* Nút xác nhận ký */}
            <div className="pt-1 flex gap-2">
              <button
                type="button"
                onClick={handleReset}
                className="bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold py-2 px-3 rounded-md transition-all active:scale-95 cursor-pointer"
              >
                Hủy / Quét lại
              </button>
              <button
                type="button"
                onClick={handleSign}
                disabled={isSigning}
                className="flex-1 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 text-white font-bold py-2 px-4 rounded-md shadow-xs transition-all active:scale-95 flex items-center justify-center gap-2 cursor-pointer"
              >
                {isSigning ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin text-white" />
                    Đang xác thực và tạo chữ ký số...
                  </>
                ) : (
                  <>
                    <ShieldCheck className="w-4 h-4 text-emerald-300 animate-bounce" />
                    Ký Số Bảo Mật Bằng Token CA
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        {/* Bước 3: Đã ký thành công */}
        {signed && (
          <div className="space-y-3 animate-fadeIn text-center py-2">
            <div className="inline-flex items-center justify-center w-12 h-12 rounded-full bg-emerald-100 text-emerald-600 mb-1">
              <ShieldCheck className="w-7 h-7" />
            </div>
            
            <div className="space-y-1">
              <h5 className="font-bold text-emerald-800 text-xs uppercase tracking-wide">
                Ký Số Thành Công!
              </h5>
              <p className="text-slate-600 text-[10.5px]">
                Báo cáo khảo sát đã được ký điện tử an toàn bằng thiết bị USB Token.
              </p>
            </div>

            {/* Hiển thị mộc đỏ điện tử của Doanh nghiệp */}
            <div className="border border-rose-200 bg-rose-50/15 rounded-md p-2 mx-auto max-w-sm flex flex-col items-center">
              <span className="text-[9px] font-bold text-rose-700 uppercase mb-1">Ảnh dấu mộc Chữ ký số phát sinh:</span>
              <img
                src={generateRedSealImage(
                  enterpriseName,
                  mst,
                  representative,
                  CA_PROVIDERS.find(c => c.id === provider)?.name || "VNPT-CA",
                  serialNumber,
                  signedTime
                )}
                alt="Chữ ký số doanh nghiệp"
                className="max-h-24 object-contain border border-dashed border-rose-300 rounded"
              />
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded p-2 text-left space-y-1 font-mono text-[9px] text-slate-600">
              <p><strong className="text-slate-800">Hash SHA-256:</strong> {shaHash}</p>
              <p><strong className="text-slate-800">Thời gian ký:</strong> {signedTime}</p>
              <p><strong className="text-slate-800">Token Serial:</strong> {serialNumber}</p>
            </div>

            <button
              type="button"
              onClick={handleReset}
              className="text-[10px] text-red-600 hover:text-red-700 underline font-bold focus:outline-none flex items-center gap-1 mx-auto pt-1 cursor-pointer"
            >
              <RefreshCw className="w-2.5 h-2.5" />
              Xóa chữ ký số để ký lại
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
