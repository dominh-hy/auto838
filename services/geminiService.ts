
import { GoogleGenAI, Type } from "@google/genai";
import { ExtractedData } from "../types";

// Sử dụng API_KEY từ biến môi trường theo hướng dẫn
const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

export async function extractDataFromImages(cccdFront: string, cccdBack: string, gdkd: string): Promise<ExtractedData> {
  const prompt = `
    Bạn là một trợ lý AI chuyên nghiệp, chuyên gia xử lý hồ sơ và hợp đồng viễn thông - công nghệ. 
    Nhiệm vụ của bạn là trích xuất dữ liệu cực kỳ chính xác từ 3 hình ảnh: CCCD mặt trước, CCCD mặt sau và Giấy đăng ký kinh doanh (GDKD).

    YÊU CẦU ĐẶC BIỆT VỀ NGÔN NGỮ:
    - Đảm bảo tiếng Việt đúng chính tả 100%.
    - Giữ đúng định dạng VIẾT HOA của Tên Công ty/Hộ kinh doanh.
    - Chuẩn hóa các danh từ riêng (Họ tên, Địa chỉ) theo quy tắc viết hoa đầu từ.

    YÊU CẦU TRÍCH XUẤT CHI TIẾT:
    1. CCCD Mặt trước: 
       - "fullname": Họ và tên (ví dụ: Nguyễn Văn A).
       - "cccdNumber": Số CCCD (12 chữ số).
       - "permanentAddress": Nơi thường trú (ghi đầy đủ).
    2. CCCD Mặt sau: 
       - "issueDate": Ngày cấp (dd/mm/yyyy).
       - "issuePlace": Nơi cấp (Phân loại đúng: "BỘ CÔNG AN" hoặc "Cục quản lý HC và TTXH").
    3. Giấy ĐKKD:
       - "businessName": Tên hộ kinh doanh/doanh nghiệp. PHẢI GIỮ NGUYÊN 100% ĐỊNH DẠNG (VIẾT HOA).
       - "registrationNumber": Số GĐKKD hoặc Mã số doanh nghiệp.
       - "registrationDate": ƯU TIÊN lấy ngày tại dòng "Đăng ký lần đầu". Nếu không có, lấy ngày cấp gần nhất.
       - "address": Địa chỉ trụ sở chính.
       - "phone": Số điện thoại liên lạc.
       - "email": Email (nếu có, nếu không trả về chuỗi rỗng).
       - "taxCode": Mã số thuế.

    LƯU Ý NGHIỆP VỤ:
    - Nếu thông tin bị mờ, hãy cố gắng suy luận dựa trên ngữ cảnh pháp lý nhưng không được bịa đặt.
  `;

  const response = await ai.models.generateContent({
    model: "gemini-3-flash-preview",
    contents: {
      parts: [
        { text: prompt },
        { inlineData: { mimeType: "image/jpeg", data: cccdFront } },
        { inlineData: { mimeType: "image/jpeg", data: cccdBack } },
        { inlineData: { mimeType: "image/jpeg", data: gdkd } }
      ]
    },
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          fullname: { type: Type.STRING },
          cccdNumber: { type: Type.STRING },
          issueDate: { type: Type.STRING },
          issuePlace: { type: Type.STRING },
          permanentAddress: { type: Type.STRING },
          businessName: { type: Type.STRING },
          registrationNumber: { type: Type.STRING },
          registrationDate: { type: Type.STRING },
          address: { type: Type.STRING },
          phone: { type: Type.STRING },
          representativeRole: { type: Type.STRING },
          email: { type: Type.STRING },
          taxCode: { type: Type.STRING }
        },
        required: ["fullname", "cccdNumber", "businessName", "address"]
      }
    }
  });

  return JSON.parse(response.text || "{}") as ExtractedData;
}

export async function getHighlightMapping(snippets: string[]): Promise<string[]> {
  const prompt = `
    Bạn là chuyên gia pháp lý AI chuyên nghiệp. Hãy gán nhãn trường dữ liệu cho các vùng [TÔ VÀNG] trong file Word dựa trên ngữ cảnh xung quanh.
    
    DANH SÁCH NHÃN HỢP LỆ:
    - contractNumber: Số hợp đồng
    - businessName: Tên công ty/Bên A
    - address: Địa chỉ trụ sở
    - phone: Số điện thoại
    - taxCode: Mã số thuế (Bên A dùng số CCCD làm MST)
    - fullname: Họ và tên người đại diện/Khách hàng
    - representativeRole: Chức vụ (Chủ hộ/Giám đốc)
    - email: Email liên hệ
    - issueDate: Ngày cấp CCCD
    - issuePlace: Nơi cấp CCCD
    - permanentAddress: Hộ khẩu thường trú
    - registrationNumber: Số GĐKKD/Quyết định thành lập
    - registrationDate: Ngày cấp GĐKKD (Ngày đăng ký lần đầu)
    - currentDay: Số Ngày hiện tại (2 chữ số)
    - currentMonth: Số Tháng hiện tại (2 chữ số)
    - currentYear: Số Năm hiện tại (4 chữ số)
    - cccdNumber: Số CCCD

    QUY TẮC ÁNH XẠ ĐẶC THÙ (RẤT QUAN TRỌNG):
    1. Phiếu yêu cầu: Cụm "Số CCCD/chứng minh thư/hộ chiếu:" -> nhãn "cccdNumber".
    2. Phiếu yêu cầu: Dòng kết thúc "Hưng Yên, ngày... tháng... năm..." -> các nhãn "currentDay", "currentMonth", "currentYear".
    3. Phiếu yêu cầu: Mục II "Cấp ngày*" -> nhãn "registrationDate".
    4. Hợp đồng & Phụ lục: Các cụm "ngày... tháng... năm..." liên quan đến thời điểm lập -> "currentDay", "currentMonth", "currentYear".
    5. Mục BÊN A: Mã số thuế -> "taxCode" (Giá trị này sẽ là số CCCD).

    YÊU CẦU: Chỉ trả về mảng chuỗi các nhãn tương ứng với danh sách snippet dưới đây.
    DANH SÁCH SNIPPETS CẦN PHÂN LOẠI:
    ${snippets.map((s, i) => `${i}. "${s}"`).join('\n')}
  `;

  const response = await ai.models.generateContent({
    model: "gemini-3-flash-preview",
    contents: prompt,
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: { type: Type.STRING }
      }
    }
  });

  try {
    return JSON.parse(response.text || "[]");
  } catch (e) {
    return [];
  }
}
