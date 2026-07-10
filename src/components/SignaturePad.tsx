import React, { useRef, useState, useEffect } from "react";
import { Eraser, Check, Paintbrush } from "lucide-react";

interface SignaturePadProps {
  onSave: (dataUrl: string) => void;
  onClear?: () => void;
  defaultValue?: string;
  penColor?: string;
  id?: string;
  height?: number;
}

export const SignaturePad: React.FC<SignaturePadProps> = ({
  onSave,
  onClear,
  defaultValue = "",
  penColor = "#0f172a", // slate-900
  id,
  height = 160,
}) => {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [hasDrawn, setHasDrawn] = useState(false);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;

    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    // Set canvas resolution to match its display size for sharp rendering
    const rect = canvas.getBoundingClientRect();
    canvas.width = rect.width * 2;
    canvas.height = rect.height * 2;
    ctx.scale(2, 2);

    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.lineWidth = 2.5;
    ctx.strokeStyle = penColor;

    // Draw background white
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, rect.width, rect.height);

    if (defaultValue) {
      const img = new Image();
      img.onload = () => {
        ctx.drawImage(img, 0, 0, rect.width, rect.height);
        setHasDrawn(true);
      };
      img.src = defaultValue;
    }
  }, [defaultValue, penColor]);

  const getCoordinates = (e: React.MouseEvent | React.TouchEvent) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };

    const rect = canvas.getBoundingClientRect();
    
    // Check if it's touch or mouse event
    if ("touches" in e) {
      if (e.touches.length === 0) return { x: 0, y: 0 };
      const touch = e.touches[0];
      return {
        x: touch.clientX - rect.left,
        y: touch.clientY - rect.top,
      };
    } else {
      return {
        x: e.clientX - rect.left,
        y: e.clientY - rect.top,
      };
    }
  };

  const startDrawing = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    const coords = getCoordinates(e);
    const canvas = canvasRef.current;
    if (!canvas) return;

    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    ctx.beginPath();
    ctx.moveTo(coords.x, coords.y);
    setIsDrawing(true);
  };

  const draw = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    if (!isDrawing) return;
    e.preventDefault();

    const coords = getCoordinates(e);
    const canvas = canvasRef.current;
    if (!canvas) return;

    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    ctx.lineTo(coords.x, coords.y);
    ctx.stroke();
    setHasDrawn(true);
  };

  const stopDrawing = () => {
    if (!isDrawing) return;
    setIsDrawing(false);
    saveSignature();
  };

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;

    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    const rect = canvas.getBoundingClientRect();
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, rect.width, rect.height);
    
    setHasDrawn(false);
    onClear?.();
  };

  const saveSignature = () => {
    const canvas = canvasRef.current;
    if (!canvas || !hasDrawn) return;

    // Output at normal resolution
    const tempCanvas = document.createElement("canvas");
    const rect = canvas.getBoundingClientRect();
    tempCanvas.width = rect.width;
    tempCanvas.height = rect.height;
    const tempCtx = tempCanvas.getContext("2d");
    
    if (tempCtx) {
      tempCtx.drawImage(canvas, 0, 0, rect.width, rect.height);
      const dataUrl = tempCanvas.toDataURL("image/png");
      onSave(dataUrl);
    }
  };

  return (
    <div className="space-y-2">
      <div className="relative border border-slate-300 rounded-lg overflow-hidden shadow-xs bg-white">
        <div className="absolute top-2 left-2 bg-slate-900/80 text-white text-[9px] font-mono px-2 py-0.5 rounded-full flex items-center gap-1 z-10 pointer-events-none uppercase">
          <Paintbrush className="w-2.5 h-2.5 animate-pulse text-indigo-300" />
          Ký trực tiếp lên bảng vẽ bên dưới
        </div>
        
        <canvas
          ref={canvasRef}
          onMouseDown={startDrawing}
          onMouseMove={draw}
          onMouseUp={stopDrawing}
          onMouseLeave={stopDrawing}
          onTouchStart={startDrawing}
          onTouchMove={draw}
          onTouchEnd={stopDrawing}
          style={{ height: height }}
          className="w-full cursor-crosshair touch-none bg-white block"
        />

        {hasDrawn && (
          <div className="absolute bottom-2 right-2 flex items-center gap-1.5 z-10">
            <button
              type="button"
              onClick={clearCanvas}
              className="bg-white hover:bg-red-50 text-red-600 border border-red-200 px-2.5 py-1 rounded-md text-[10.5px] font-bold flex items-center gap-1 shadow-2xs transition active:scale-95 cursor-pointer"
            >
              <Eraser className="w-3 h-3" />
              Xóa vẽ lại
            </button>
          </div>
        )}
      </div>
      <p className="text-[10px] text-slate-500 italic">
        Sử dụng ngón tay (trên điện thoại/máy tính bảng) hoặc giữ chuột trái (trên máy tính) để vẽ nét ký tay.
      </p>
    </div>
  );
};
