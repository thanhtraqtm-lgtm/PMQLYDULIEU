import React from "react";
import { PhanBoHoXa } from "./PhanBoHoXa";

interface SamplingSelectionProps {
  [key: string]: any;
}

export const SamplingSelection: React.FC<SamplingSelectionProps> = () => {
  return (
    <div className="w-full">
      {/* Ứng dụng phân bổ hộ điều tra cấp xã (Mỗi sheet là 1 địa bàn) */}
      <PhanBoHoXa />
    </div>
  );
};

export default SamplingSelection;
