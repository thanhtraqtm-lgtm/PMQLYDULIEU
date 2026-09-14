import React, { useRef, useState, useEffect } from "react";
import { RotateCcw, Check, PenTool } from "lucide-react";

interface SignaturePadProps {
  id?: string;
  onSave: (dataUrl: string) => void;
  onClear?: () => void;
  height?: number;
  width?: number;
}

export const SignaturePad: React.FC<SignaturePadProps> = ({
  id = "signature-pad",
  onSave,
  onClear,
  height = 110,
}) => {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [hasDrawn, setHasDrawn] = useState(false);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    // Set high-DPI scaling
    const rect = canvas.getBoundingClientRect();
    canvas.width = rect.width * 2;
    canvas.height = height * 2;
    ctx.scale(2, 2);
    ctx.strokeStyle = "#1e293b";
    ctx.lineWidth = 2.5;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
  }, [height]);

  const getCoordinates = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    const rect = canvas.getBoundingClientRect();

    if ("touches" in e) {
      const touch = e.touches[0];
      return {
        x: touch.clientX - rect.left,
        y: touch.clientY - rect.top,
      };
    }
    return {
      x: e.clientX - rect.left,
      y: e.clientY - rect.top,
    };
  };

  const startDrawing = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    const { x, y } = getCoordinates(e);
    ctx.beginPath();
    ctx.moveTo(x, y);
    setIsDrawing(true);
    setHasDrawn(true);
  };

  const draw = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    if (!isDrawing) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    const { x, y } = getCoordinates(e);
    ctx.lineTo(x, y);
    ctx.stroke();
  };

  const stopDrawing = () => {
    if (!isDrawing) return;
    setIsDrawing(false);
    const canvas = canvasRef.current;
    if (canvas) {
      onSave(canvas.toDataURL("image/png"));
    }
  };

  const handleClear = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasDrawn(false);
    onSave("");
    onClear?.();
  };

  return (
    <div className="space-y-1.5" id={id}>
      <div className="relative border border-slate-300 rounded-lg overflow-hidden bg-white shadow-2xs">
        <canvas
          ref={canvasRef}
          onMouseDown={startDrawing}
          onMouseMove={draw}
          onMouseUp={stopDrawing}
          onMouseLeave={stopDrawing}
          onTouchStart={startDrawing}
          onTouchMove={draw}
          onTouchEnd={stopDrawing}
          style={{ height: `${height}px`, width: "100%", touchAction: "none" }}
          className="cursor-crosshair block w-full"
        />

        {!hasDrawn && (
          <div className="absolute inset-0 pointer-events-none flex items-center justify-center text-slate-300 text-[11px] select-none">
            Ký tay tại đây...
          </div>
        )}
      </div>

      <div className="flex items-center justify-between text-[11px]">
        <button
          type="button"
          onClick={handleClear}
          className="inline-flex items-center gap-1 text-slate-500 hover:text-rose-600 transition-colors cursor-pointer py-0.5"
        >
          <RotateCcw className="w-3 h-3" />
          Xóa ký lại
        </button>
        {hasDrawn && (
          <span className="text-emerald-600 font-semibold inline-flex items-center gap-0.5">
            <Check className="w-3 h-3" /> Đã ghi nhận chữ ký
          </span>
        )}
      </div>
    </div>
  );
};
export default SignaturePad;
