
import { ExtractedData, ContractData } from "../types";

export function getInitials(name: string): string {
  if (!name) return "";
  return name
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .split(/\s+/)
    .map(word => word[0])
    .join("")
    .toUpperCase();
}

export function generateContractData(raw: ExtractedData): ContractData {
  const initials = getInitials(raw.fullname);
  const contractNumber = `${raw.cccdNumber}/HĐ-MOBIFONE/MBF-${initials}`;
  
  // Lấy ngày tháng hiện tại
  const now = new Date();
  const day = String(now.getDate()).padStart(2, '0');
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const year = String(now.getFullYear());
  
  const formattedDate = `ngày ${day} tháng ${month} năm ${year}`;

  // Logic chức vụ tự động theo yêu cầu:
  // Hộ kinh doanh -> Chủ hộ
  // Công ty (TNHH, CP, Doanh nghiệp) -> Giám đốc
  let role = raw.representativeRole || "";
  const bName = (raw.businessName || "").toUpperCase();
  
  if (bName.includes("HỘ KINH DOANH") || bName.includes("HKD")) {
    role = "Chủ hộ";
  } else if (
    bName.includes("CÔNG TY") || 
    bName.includes("TNHH") || 
    bName.includes("CỔ PHẦN") || 
    bName.includes("CP") ||
    bName.includes("DOANH NGHIỆP")
  ) {
    role = "Giám đốc";
  } else {
    // Mặc định nếu không xác định rõ
    role = "Đại diện";
  }
  
  return {
    ...raw,
    id: Math.random().toString(36).substr(2, 9),
    contractNumber,
    representativeRole: role,
    // Bên A dùng số CCCD làm Mã số thuế
    taxCode: raw.cccdNumber,
    email: raw.email && raw.email.trim() !== "" ? raw.email : "yenmy602@gmail.com",
    creationDate: formattedDate,
    currentDay: day,
    currentMonth: month,
    currentYear: year,
    status: 'pending'
  };
}
