
export interface ExtractedData {
  fullname: string;
  cccdNumber: string;
  issueDate: string;        // Ngày cấp CCCD
  issuePlace: string;       // Nơi cấp CCCD
  permanentAddress: string; // Hộ khẩu thường trú
  businessName: string;
  registrationNumber: string; // Số GĐKKD / Quyết định thành lập
  registrationDate: string;   // Ngày cấp GĐKKD
  address: string;
  phone: string;
  representativeRole: string;
  email: string;
  taxCode: string;
}

export interface ContractData extends ExtractedData {
  id: string;
  contractNumber: string;
  creationDate: string;
  currentDay: string;   // Ngày hiện tại (2 chữ số)
  currentMonth: string; // Tháng hiện tại (2 chữ số)
  currentYear: string;  // Năm hiện tại (4 chữ số)
  status: 'pending' | 'validated' | 'processing' | 'completed' | 'error';
}

export interface ValidationError {
  field: string;
  message: string;
}

export enum AppStep {
  DASHBOARD = 1,
  UPLOAD_IMAGES = 2,
  UPLOAD_WORD = 3,
  EXPORT_ALL = 4
}
