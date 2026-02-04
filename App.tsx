
import React, { useState, useEffect } from 'react';
import { 
  FileText, 
  Upload, 
  CheckCircle, 
  AlertCircle, 
  Plus, 
  Download, 
  Edit3,
  Image as ImageIcon,
  FileDown,
  Loader2,
  Trash2,
  Files,
  X,
  ChevronRight,
  UserCheck,
  Info,
  Eye,
  ArrowRight,
  User,
  Settings,
  Printer,
  FileSearch,
  Share2,
  Send,
  Copy,
  Mail
} from 'lucide-react';
import { AppStep, ContractData, ValidationError } from './types';
import { extractDataFromImages, getHighlightMapping } from './services/geminiService';
import { generateContractData } from './services/ContractService';
import PizZip from 'pizzip';

declare const mammoth: any;
declare const html2pdf: any;

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const fileToBase64 = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve((reader.result as string).split(',')[1]);
    reader.onerror = error => reject(error);
  });
};

const App: React.FC = () => {
  const [step, setStep] = useState<AppStep>(AppStep.DASHBOARD);
  const [loading, setLoading] = useState(false);
  const [loadingMessage, setLoadingMessage] = useState("");
  
  const [contracts, setContracts] = useState<ContractData[]>([]);
  const [selectedContractId, setSelectedContractId] = useState<string | null>(null);
  
  const [newCccdFront, setNewCccdFront] = useState<File | null>(null);
  const [newCccdBack, setNewCccdBack] = useState<File | null>(null);
  const [newGdkd, setNewGdkd] = useState<File | null>(null);
  
  const [contractBuffer, setContractBuffer] = useState<ArrayBuffer | null>(null);
  const [contractFileName, setContractFileName] = useState<string>("");
  const [requestBuffer, setRequestBuffer] = useState<ArrayBuffer | null>(null);
  const [requestFileName, setRequestFileName] = useState<string>("");
  
  const [previewHtml, setPreviewHtml] = useState<string>("");
  const [activePreviewType, setActivePreviewType] = useState<'contract' | 'request'>('contract');

  const activeContract = contracts.find(c => c.id === selectedContractId);

  // CSS nội sinh chuyên dụng cho các tác vụ xuất (Print/PDF)
  const internalDocStyles = `
    <style>
      @page { size: A4; margin: 20mm 15mm 20mm 30mm; }
      body { margin: 0; padding: 0; background: #fff; }
      .document-content {
        font-family: 'Times New Roman', Times, serif !important;
        line-height: 1.4 !important;
        color: #000 !important;
        text-align: justify !important;
        font-size: 13pt !important;
        padding: 0;
        margin: 0;
      }
      .document-content p { margin: 0 0 10pt 0 !important; text-align: justify !important; }
      .document-content h1, .document-content h2, .document-content h3 { 
        text-align: center !important; 
        font-weight: bold !important; 
        text-transform: uppercase !important;
        margin: 12pt 0 !important;
        font-size: 14pt !important;
      }
      .document-content table { 
        width: 100% !important; 
        border-collapse: collapse !important; 
        margin: 10pt 0 !important;
        border: 1px solid black !important;
      }
      .document-content td, .document-content th { 
        border: 1px solid black !important; 
        padding: 6pt 8pt !important; 
        vertical-align: top !important;
        font-size: 12pt !important;
      }
    </style>
  `;

  useEffect(() => {
    const buffer = activePreviewType === 'contract' ? contractBuffer : requestBuffer;
    if (buffer) {
      mammoth.convertToHtml({ arrayBuffer: buffer })
        .then((result: any) => setPreviewHtml(result.value))
        .catch(() => setPreviewHtml("<p class='text-red-500'>Lỗi hiển thị nội dung mẫu.</p>"));
    } else {
      setPreviewHtml("");
    }
  }, [contractBuffer, requestBuffer, activePreviewType]);

  const handleDownloadSample = (type: 'contract' | 'request') => {
    setLoading(true);
    setLoadingMessage(`Đang tạo mẫu ${type === 'contract' ? 'Hợp đồng' : 'Phiếu yêu cầu'}...`);
    try {
      const zip = new PizZip();
      const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
      const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
      
      let docXml = "";
      if (type === 'contract') {
        docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>HỢP ĐỒNG DỊCH VỤ VIỄN THÔNG</w:t></w:r></w:p><w:p><w:r><w:t>Bên A: </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>businessName</w:t></w:r></w:p><w:p><w:r><w:t>Số hợp đồng: </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>contractNumber</w:t></w:r></w:p><w:p><w:r><w:t>Người đại diện: </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>fullname</w:t></w:r></w:p><w:p><w:r><w:t>Mã số thuế: </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>taxCode</w:t></w:r></w:p><w:p><w:r><w:t>Hưng Yên, </w:t></w:r><w:r><w:t>ngày </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>currentDay</w:t></w:r><w:r><w:t> tháng </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>currentMonth</w:t></w:r><w:r><w:t> năm </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>currentYear</w:t></w:r></w:p></w:body></w:document>`;
      } else {
        docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>PHIẾU YÊU CẦU DỊCH VỤ</w:t></w:r></w:p><w:p><w:r><w:t>Mục II: Thông tin ĐKKD</w:t></w:r></w:p><w:p><w:r><w:t>Cấp ngày*: </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>registrationDate</w:t></w:r></w:p><w:p><w:r><w:t>Số CCCD/chứng minh thư/hộ chiếu: </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>cccdNumber</w:t></w:r></w:p><w:p><w:r><w:t>Hưng Yên, ngày </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>currentDay</w:t></w:r><w:r><w:t> tháng </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>currentMonth</w:t></w:r><w:r><w:t> năm </w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>currentYear</w:t></w:r></w:p></w:body></w:document>`;
      }

      zip.file("[Content_Types].xml", contentTypes);
      zip.file("_rels/.rels", rels);
      zip.file("word/document.xml", docXml);
      
      const blob = zip.generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `MAU_${type === 'contract' ? 'HOP_DONG' : 'PHIEU_YEU_CAU'}.docx`;
      link.click();
    } catch (e) {
      alert("Lỗi tạo file mẫu.");
    } finally {
      setLoading(false);
    }
  };

  const handleOcr = async () => {
    if (!newCccdFront || !newCccdBack || !newGdkd) return;
    setLoading(true);
    setLoadingMessage("AI đang đọc dữ liệu và xử lý tiếng Việt...");
    try {
      const frontB64 = await fileToBase64(newCccdFront);
      const backB64 = await fileToBase64(newCccdBack);
      const gdkdB64 = await fileToBase64(newGdkd);
      const extracted = await extractDataFromImages(frontB64, backB64, gdkdB64);
      const processed = generateContractData(extracted);
      setContracts(prev => [...prev, processed]);
      setNewCccdFront(null); setNewCccdBack(null); setNewGdkd(null);
      setStep(AppStep.DASHBOARD);
    } catch (error) {
      alert("Lỗi trích xuất hồ sơ.");
    } finally {
      setLoading(false);
    }
  };

  const processWordInternal = async (contract: ContractData, type: 'contract' | 'request') => {
    const buffer = type === 'contract' ? contractBuffer : requestBuffer;
    if (!buffer) throw new Error("Chưa có mẫu file");

    const zip = new PizZip(buffer);
    const docXmlStr = zip.file("word/document.xml")?.asText();
    if (!docXmlStr) throw new Error("Cấu trúc file lỗi");

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(docXmlStr, "application/xml");
    const runs = Array.from(xmlDoc.getElementsByTagNameNS(WORD_NS, "r"));
    
    const fieldBlocks: Element[][] = [];
    let currentBlock: Element[] = [];
    runs.forEach((run) => {
      const hasHighlight = Array.from(run.getElementsByTagNameNS(WORD_NS, "highlight")).some(h => h.getAttributeNS(WORD_NS, "val") === "yellow");
      if (hasHighlight) currentBlock.push(run);
      else if (currentBlock.length > 0) { fieldBlocks.push([...currentBlock]); currentBlock = []; }
    });
    if (currentBlock.length > 0) fieldBlocks.push(currentBlock);

    if (fieldBlocks.length > 0) {
      const snippets = fieldBlocks.map(block => {
        const prev = block[0].previousElementSibling?.textContent || "";
        const text = block.map(r => r.getElementsByTagNameNS(WORD_NS, "t")[0]?.textContent || "").join("");
        return `${prev} [TÔ VÀNG: ${text}]`;
      });

      const mapping = await getHighlightMapping(snippets);

      fieldBlocks.forEach((block, index) => {
        const field = mapping[index] as keyof ContractData;
        if (field && contract[field]) {
          const firstT = block[0].getElementsByTagNameNS(WORD_NS, "t")[0];
          if (firstT) firstT.textContent = String(contract[field]);
          for (let i = 1; i < block.length; i++) {
            const otherT = block[i].getElementsByTagNameNS(WORD_NS, "t")[0];
            if (otherT) otherT.textContent = "";
          }
          block.forEach(r => {
            const h = r.getElementsByTagNameNS(WORD_NS, "highlight")[0];
            if (h) h.parentNode?.removeChild(h);
          });
        }
      });
    }

    const updatedXml = new XMLSerializer().serializeToString(xmlDoc);
    zip.file("word/document.xml", updatedXml);
    return zip.generate({ type: "arraybuffer" });
  };

  const handleProcessWord = async (contract: ContractData, type: 'contract' | 'request') => {
    const buffer = type === 'contract' ? contractBuffer : requestBuffer;
    if (!buffer) { setActivePreviewType(type); setStep(AppStep.UPLOAD_WORD); return; }
    setLoading(true);
    setLoadingMessage(`Đang xuất file Word cho ${contract.businessName}...`);
    try {
      const resultBuffer = await processWordInternal(contract, type);
      const blob = new Blob([resultBuffer], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${type === 'contract' ? 'HOP_DONG' : 'PHIEU_YEU_CAU'}_${contract.businessName.replace(/\s+/g, '_')}.docx`;
      link.click();
      updateContract(contract.id, { status: 'completed' });
    } catch (e: any) { alert(e.message || "Lỗi xử lý file."); } finally { setLoading(false); }
  };

  const handleExportPdf = async (contract: ContractData, type: 'contract' | 'request') => {
    const buffer = type === 'contract' ? contractBuffer : requestBuffer;
    if (!buffer) { setActivePreviewType(type); setStep(AppStep.UPLOAD_WORD); return; }
    setLoading(true);
    setLoadingMessage(`Đang chuyển đổi và xuất file PDF...`);
    try {
      const resultBuffer = await processWordInternal(contract, type);
      const result = await mammoth.convertToHtml({ arrayBuffer: resultBuffer });
      const html = result.value;
      const element = document.createElement('div');
      element.innerHTML = `${internalDocStyles}<div class="document-content">${html}</div>`;
      const opt = {
        margin: [20, 15, 20, 30],
        filename: `${type === 'contract' ? 'HOP_DONG' : 'PHIEU_YEU_CAU'}_${contract.businessName.replace(/\s+/g, '_')}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 3, useCORS: true, letterRendering: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
      };
      await html2pdf().set(opt).from(element).save();
      updateContract(contract.id, { status: 'completed' });
    } catch (e: any) { alert("Lỗi xuất PDF: " + e.message); } finally { setLoading(false); }
  };

  const handlePrint = async (contract: ContractData, type: 'contract' | 'request') => {
    const buffer = type === 'contract' ? contractBuffer : requestBuffer;
    if (!buffer) { setActivePreviewType(type); setStep(AppStep.UPLOAD_WORD); return; }
    setLoading(true);
    setLoadingMessage("Đang chuẩn bị bản in chuyên nghiệp...");
    try {
      const resultBuffer = await processWordInternal(contract, type);
      const result = await mammoth.convertToHtml({ arrayBuffer: resultBuffer });
      const html = result.value;
      const printWindow = window.open('', '_blank');
      if (printWindow) {
        printWindow.document.write(`
          <!DOCTYPE html><html><head><title>In ${type === 'contract' ? 'Hợp đồng' : 'Phiếu yêu cầu'}</title><meta charset="UTF-8">${internalDocStyles}</head><body><div class="document-content">${html}</div><script>window.onload = function() { setTimeout(() => { window.print(); window.onafterprint = function() { window.close(); }; }, 800); }</script></body></html>
        `);
        printWindow.document.close();
      }
    } catch (e: any) { alert("Lỗi chuẩn bị bản in: " + e.message); } finally { setLoading(false); }
  };

  const handleShareFileAction = async (contract: ContractData, type: 'contract' | 'request', format: 'docx' | 'pdf') => {
    setLoading(true);
    setLoadingMessage(`Đang chuẩn bị file để chia sẻ qua Zalo/Gmail...`);
    try {
      const resultBuffer = await processWordInternal(contract, type);
      let blob: Blob;
      let filename: string;
      let mimeType: string;

      if (format === 'docx') {
        blob = new Blob([resultBuffer], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        filename = `${type === 'contract' ? 'HĐ' : 'PYC'}_${contract.businessName.replace(/\s+/g, '_')}.docx`;
        mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      } else {
        const result = await mammoth.convertToHtml({ arrayBuffer: resultBuffer });
        const html = result.value;
        const element = document.createElement('div');
        element.innerHTML = `${internalDocStyles}<div class="document-content">${html}</div>`;
        blob = await html2pdf().set({
            margin: [20, 15, 20, 30],
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 3, useCORS: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        }).from(element).outputPdf('blob');
        filename = `${type === 'contract' ? 'HĐ' : 'PYC'}_${contract.businessName.replace(/\s+/g, '_')}.pdf`;
        mimeType = "application/pdf";
      }

      const file = new File([blob], filename, { type: mimeType });

      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({
          files: [file],
          title: `Hồ sơ Mobifone: ${contract.businessName}`,
          text: `Gửi file ${type === 'contract' ? 'Hợp đồng' : 'Phiếu yêu cầu'} của khách hàng ${contract.fullname}.`,
          url: window.location.href
        });
      } else {
        // Fallback: Copy link and info
        const shareText = `Hồ sơ: ${contract.businessName}\nNgười đại diện: ${contract.fullname}\nSố HĐ: ${contract.contractNumber}\nLink truy cập: ${window.location.href}`;
        await navigator.clipboard.writeText(shareText);
        alert("Trình duyệt không hỗ trợ gửi file trực tiếp. Đã sao chép link và thông tin hồ sơ vào bộ nhớ tạm!");
      }
    } catch (e: any) {
      if (e.name !== 'AbortError') alert("Lỗi chia sẻ: " + e.message);
    } finally { setLoading(false); }
  };

  const handleQuickShareInfo = async (contract: ContractData) => {
    const text = `HỒ SƠ KHÁCH HÀNG MOBIFONE\n--------------------------\n- Đơn vị: ${contract.businessName}\n- Người đại diện: ${contract.fullname}\n- CCCD/MST: ${contract.cccdNumber}\n- Số HĐ: ${contract.contractNumber}\n- Ngày lập: ${contract.creationDate}`;
    try {
      if (navigator.share) {
        await navigator.share({ title: 'Thông tin hồ sơ', text: text, url: window.location.href });
      } else {
        await navigator.clipboard.writeText(text);
        alert("Đã sao chép thông tin vào bộ nhớ tạm!");
      }
    } catch (e) {}
  };

  const updateContract = (id: string, updates: Partial<ContractData>) => {
    setContracts(prev => prev.map(c => c.id === id ? { ...c, ...updates } : c));
  };

  const deleteContract = (id: string) => {
    if (confirm("Xác nhận xóa hồ sơ này?")) {
      setContracts(prev => prev.filter(c => c.id !== id));
      if (selectedContractId === id) setSelectedContractId(null);
    }
  };

  const ImageUploadBox = ({ label, file, setFile, icon: Icon }: { label: string, file: File | null, setFile: (f: File | null) => void, icon: any }) => (
    <div className="space-y-3">
      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2">
        <Icon size={14} className="text-blue-500"/> {label}
      </label>
      <div className="relative group overflow-hidden rounded-3xl h-40">
        <input type="file" className="absolute inset-0 opacity-0 cursor-pointer z-20" accept="image/*" onChange={(e) => setFile(e.target.files?.[0] || null)} />
        <div className={`absolute inset-0 border-2 border-dashed transition-all flex flex-col items-center justify-center p-4 ${file ? 'bg-white border-green-400' : 'bg-slate-50 border-slate-200 group-hover:border-blue-400'}`}>
          {file ? (
            <>
              <img src={URL.createObjectURL(file)} className="absolute inset-0 w-full h-full object-cover opacity-20" alt="preview" />
              <div className="relative z-10 flex flex-col items-center">
                <div className="bg-green-500 text-white p-2 rounded-full mb-2"><CheckCircle size={24} /></div>
                <span className="text-xs font-bold text-green-700 truncate max-w-[150px]">{file.name}</span>
              </div>
            </>
          ) : (
            <>
              <Icon className="text-slate-300 mb-2" size={32} />
              <p className="text-[10px] font-bold text-slate-400 text-center uppercase">Kéo thả hoặc click</p>
            </>
          )}
        </div>
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-[#f8fafc] text-slate-900 font-sans">
      <header className="bg-white border-b border-slate-200 py-4 px-6 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="bg-blue-600 p-2.5 rounded-xl shadow-lg shadow-blue-200 text-white"><Files size={24} /></div>
            <div onClick={() => setStep(AppStep.DASHBOARD)} className="cursor-pointer">
              <h1 className="text-xl font-bold tracking-tight text-slate-900">Mobifone AI Hub</h1>
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em]">Xử lý hồ sơ & Chia sẻ thông minh</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <button onClick={() => setStep(AppStep.UPLOAD_WORD)} className="hidden sm:flex items-center gap-2 text-slate-500 hover:text-blue-600 font-bold text-xs transition-colors px-3"><Settings size={16}/> CÀI ĐẶT MẪU</button>
            <button onClick={() => setStep(AppStep.UPLOAD_IMAGES)} className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 rounded-xl font-bold text-sm transition-all shadow-lg shadow-blue-100 flex items-center gap-2"><Plus size={18} /> Thêm Hồ Sơ</button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto py-8 px-6">
        {step === AppStep.DASHBOARD && (
          <div className="grid lg:grid-cols-3 gap-8">
            <div className="lg:col-span-2 space-y-6">
              <div className="flex items-center justify-between">
                <h2 className="text-lg font-bold flex items-center gap-2">Hồ sơ khách hàng <span className="bg-blue-100 text-blue-600 px-2 py-0.5 rounded-full text-xs">{contracts.length}</span></h2>
                <div className="flex items-center gap-4">
                   <button onClick={() => handleDownloadSample('contract')} className="text-[10px] font-bold text-slate-400 hover:text-blue-600 flex items-center gap-1"><Download size={14}/> MẪU HĐ</button>
                   <button onClick={() => handleDownloadSample('request')} className="text-[10px] font-bold text-slate-400 hover:text-blue-600 flex items-center gap-1"><Download size={14}/> MẪU PHIẾU</button>
                </div>
              </div>

              {contracts.length === 0 ? (
                <div className="bg-white border-2 border-dashed border-slate-200 rounded-[32px] p-20 text-center">
                  <UserCheck size={40} className="mx-auto text-slate-200 mb-6" />
                  <h3 className="text-xl font-bold mb-2">Chưa có hồ sơ nào</h3>
                  <p className="text-slate-400 text-sm mb-8">Bắt đầu bằng cách thêm hồ sơ khách hàng từ 3 mặt giấy tờ.</p>
                  <button onClick={() => setStep(AppStep.UPLOAD_IMAGES)} className="px-10 py-4 bg-slate-900 text-white rounded-2xl font-bold shadow-xl">Thêm khách hàng</button>
                </div>
              ) : (
                <div className="space-y-3">
                  {contracts.map(contract => (
                    <div key={contract.id} onClick={() => setSelectedContractId(contract.id)} className={`group bg-white p-5 rounded-3xl border-2 transition-all cursor-pointer flex items-center justify-between ${selectedContractId === contract.id ? 'border-blue-500 shadow-xl shadow-blue-50' : 'border-transparent hover:border-slate-200 shadow-sm'}`}>
                      <div className="flex items-center gap-4">
                        <div className="w-12 h-12 rounded-2xl bg-blue-50 text-blue-600 flex items-center justify-center font-bold uppercase">{contract.fullname[0]}</div>
                        <div>
                          <h4 className="font-bold text-slate-900">{contract.businessName}</h4>
                          <p className="text-[10px] text-slate-400 font-bold uppercase">{contract.fullname} • {contract.cccdNumber}</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-2">
                        <button onClick={(e) => { e.stopPropagation(); handleQuickShareInfo(contract); }} className="p-2 text-slate-400 hover:text-blue-500 hover:bg-blue-50 rounded-full transition-all" title="Chia sẻ thông tin nhanh"><Share2 size={18}/></button>
                        <ChevronRight className={`transition-transform ${selectedContractId === contract.id ? 'rotate-90 text-blue-500' : 'text-slate-300'}`} />
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div className="lg:col-span-1">
              {activeContract ? (
                <div className="bg-white rounded-[32px] p-8 shadow-xl shadow-slate-200/50 border border-slate-100 sticky top-28 space-y-6">
                  <div className="flex items-center justify-between">
                    <h3 className="font-bold text-lg">Hồ sơ chi tiết</h3>
                    <div className="flex items-center gap-1">
                      <button onClick={() => handleQuickShareInfo(activeContract)} className="text-blue-500 p-2 hover:bg-blue-50 rounded-full transition-all" title="Chia sẻ hồ sơ"><Share2 size={18}/></button>
                      <button onClick={() => deleteContract(activeContract.id)} className="text-red-500 p-2 hover:bg-red-50 rounded-full transition-all" title="Xóa"><Trash2 size={18}/></button>
                    </div>
                  </div>
                  
                  <div className="space-y-4 max-h-[45vh] overflow-y-auto pr-2 custom-scrollbar">
                    <p className="text-[10px] font-black text-blue-500 uppercase tracking-widest border-b pb-2 flex items-center gap-2"><User size={12}/> Thông tin cá nhân</p>
                    {[
                      { label: 'Họ tên', value: activeContract.fullname, key: 'fullname' },
                      { label: 'Số CCCD', value: activeContract.cccdNumber, key: 'cccdNumber' },
                      { label: 'Hộ khẩu', value: activeContract.permanentAddress, key: 'permanentAddress' },
                    ].map(f => (
                      <div key={f.key} className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{f.label}</label>
                        <input type="text" className="w-full bg-slate-50 rounded-xl p-3 text-sm font-bold border-none focus:ring-2 focus:ring-blue-100 outline-none" value={f.value} onChange={(e) => updateContract(activeContract.id, { [f.key as keyof ContractData]: e.target.value })} />
                      </div>
                    ))}
                    <p className="text-[10px] font-black text-blue-500 uppercase tracking-widest border-b pb-2 mt-6 flex items-center gap-2"><FileText size={12}/> Thông tin BÊN A</p>
                    {[
                      { label: 'Tên Bên A', value: activeContract.businessName, key: 'businessName' },
                      { label: 'Mã số thuế', value: activeContract.taxCode, key: 'taxCode' },
                      { label: 'Chức vụ', value: activeContract.representativeRole, key: 'representativeRole' },
                      { label: 'Địa chỉ', value: activeContract.address, key: 'address' },
                    ].map(f => (
                      <div key={f.key} className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{f.label}</label>
                        <input type="text" className="w-full bg-slate-50 rounded-xl p-3 text-sm font-bold border-none focus:ring-2 focus:ring-blue-100 outline-none" value={f.value} onChange={(e) => updateContract(activeContract.id, { [f.key as keyof ContractData]: e.target.value })} />
                      </div>
                    ))}
                  </div>

                  <div className="space-y-4 pt-4 border-t">
                    <div className="space-y-3">
                      <div className="bg-slate-50 p-4 rounded-3xl space-y-3 border border-slate-100">
                        <div className="flex items-center justify-between"><span className="text-[10px] font-black uppercase text-blue-600">HỢP ĐỒNG</span>{contractBuffer && <span className="text-[8px] font-bold text-green-600">MẪU OK</span>}</div>
                        <div className="flex flex-wrap gap-2">
                          <button onClick={() => handleProcessWord(activeContract, 'contract')} className="flex-1 min-w-[80px] py-3 rounded-2xl bg-blue-600 text-white font-bold shadow-lg flex items-center justify-center gap-1.5 transition-all text-[10px] hover:bg-blue-700 disabled:bg-slate-200" disabled={!contractBuffer} title="Tải Word"><FileDown size={14} /> WORD</button>
                          <button onClick={() => handleExportPdf(activeContract, 'contract')} className="flex-1 min-w-[80px] py-3 rounded-2xl bg-red-600 text-white font-bold shadow-lg flex items-center justify-center gap-1.5 transition-all text-[10px] hover:bg-red-700 disabled:bg-slate-200" disabled={!contractBuffer} title="Tải PDF"><FileSearch size={14} /> PDF</button>
                          <button onClick={() => handleShareFileAction(activeContract, 'contract', 'docx')} className="p-3 rounded-2xl bg-white text-slate-600 border border-slate-200 font-bold shadow-sm flex items-center justify-center transition-all hover:bg-slate-50 disabled:opacity-30" disabled={!contractBuffer} title="Chia sẻ file qua Zalo/Gmail"><Send size={16} className="text-blue-500"/></button>
                          <button onClick={() => handlePrint(activeContract, 'contract')} className="p-3 rounded-2xl bg-white text-slate-600 border border-slate-200 font-bold shadow-sm flex items-center justify-center transition-all hover:bg-slate-50 disabled:opacity-30" disabled={!contractBuffer} title="In ngay"><Printer size={16} /></button>
                        </div>
                      </div>
                      
                      <div className="bg-slate-50 p-4 rounded-3xl space-y-3 border border-slate-100">
                        <div className="flex items-center justify-between"><span className="text-[10px] font-black uppercase text-slate-900">PHIẾU YÊU CẦU</span>{requestBuffer && <span className="text-[8px] font-bold text-green-600">MẪU OK</span>}</div>
                        <div className="flex flex-wrap gap-2">
                          <button onClick={() => handleProcessWord(activeContract, 'request')} className="flex-1 min-w-[80px] py-3 rounded-2xl bg-slate-900 text-white font-bold shadow-lg flex items-center justify-center gap-1.5 transition-all text-[10px] hover:bg-black disabled:bg-slate-200" disabled={!requestBuffer} title="Tải Word"><FileText size={14} /> WORD</button>
                          <button onClick={() => handleExportPdf(activeContract, 'request')} className="flex-1 min-w-[80px] py-3 rounded-2xl bg-orange-600 text-white font-bold shadow-lg flex items-center justify-center gap-1.5 transition-all text-[10px] hover:bg-orange-700 disabled:bg-slate-200" disabled={!requestBuffer} title="Tải PDF"><FileSearch size={14} /> PDF</button>
                          <button onClick={() => handleShareFileAction(activeContract, 'request', 'pdf')} className="p-3 rounded-2xl bg-white text-slate-600 border border-slate-200 font-bold shadow-sm flex items-center justify-center transition-all hover:bg-slate-50 disabled:opacity-30" disabled={!requestBuffer} title="Chia sẻ file qua Zalo/Gmail"><Send size={16} className="text-blue-500"/></button>
                          <button onClick={() => handlePrint(activeContract, 'request')} className="p-3 rounded-2xl bg-white text-slate-600 border border-slate-200 font-bold shadow-sm flex items-center justify-center transition-all hover:bg-slate-50 disabled:opacity-30" disabled={!requestBuffer} title="In ngay"><Printer size={16} /></button>
                        </div>
                      </div>
                    </div>
                    {(!contractBuffer || !requestBuffer) && <p className="text-[10px] text-red-500 font-bold text-center italic uppercase tracking-widest leading-tight mt-2 flex items-center justify-center gap-1"><AlertCircle size={12}/> Vui lòng cài đặt mẫu file trước</p>}
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded-[32px] p-20 text-center border-2 border-dashed border-slate-100">
                  <Info className="mx-auto text-slate-200 mb-4" /><p className="text-slate-400 text-sm font-medium">Chọn một hồ sơ để bắt đầu chia sẻ.</p>
                </div>
              )}
            </div>
          </div>
        )}

        {step === AppStep.UPLOAD_IMAGES && (
          <div className="max-w-4xl mx-auto animate-in fade-in slide-in-from-bottom-8 duration-500">
            <div className="bg-white p-12 rounded-[48px] shadow-2xl border border-slate-100">
              <div className="flex items-center justify-between mb-10"><h2 className="text-2xl font-black uppercase tracking-tight">Thêm hồ sơ khách hàng</h2><button onClick={() => setStep(AppStep.DASHBOARD)} className="w-12 h-12 rounded-full bg-slate-50 flex items-center justify-center text-slate-400 hover:bg-slate-100"><X size={24}/></button></div>
              <div className="grid md:grid-cols-3 gap-6 mb-12">
                <ImageUploadBox label="1. CCCD Mặt trước" file={newCccdFront} setFile={setNewCccdFront} icon={User} />
                <ImageUploadBox label="2. CCCD Mặt sau" file={newCccdBack} setFile={setNewCccdBack} icon={ImageIcon} />
                <ImageUploadBox label="3. Giấy ĐKKD" file={newGdkd} setFile={setNewGdkd} icon={FileText} />
              </div>
              <div className="flex gap-4">
                <button onClick={() => setStep(AppStep.DASHBOARD)} className="flex-1 py-5 border-2 border-slate-100 rounded-[24px] font-bold text-slate-500 hover:bg-slate-50 uppercase tracking-widest text-xs">HỦY BỎ</button>
                <button onClick={handleOcr} disabled={loading || !newCccdFront || !newCccdBack || !newGdkd} className="flex-[2] py-5 bg-blue-600 disabled:bg-slate-200 text-white rounded-[24px] font-black shadow-xl shadow-blue-100 flex items-center justify-center gap-3 transition-all hover:bg-blue-700 uppercase tracking-widest text-sm">{loading ? <Loader2 className="animate-spin" /> : <Plus />} BẮT ĐẦU TRÍCH XUẤT</button>
              </div>
            </div>
          </div>
        )}

        {step === AppStep.UPLOAD_WORD && (
          <div className="max-w-6xl mx-auto animate-in zoom-in duration-300">
            <div className="bg-white p-12 rounded-[48px] shadow-2xl border border-slate-100">
              <div className="flex items-center justify-between mb-8"><div><h2 className="text-2xl font-black">Cấu hình Mẫu File</h2><p className="text-slate-400 text-sm font-medium">Tải lên mẫu file Word có các vùng tô vàng.</p></div><button onClick={() => setStep(AppStep.DASHBOARD)} className="text-slate-400 hover:text-slate-600"><X size={24}/></button></div>
              <div className="grid lg:grid-cols-2 gap-12">
                <div className="space-y-8">
                  <div className={`p-6 rounded-[32px] border-2 transition-all cursor-pointer ${activePreviewType === 'contract' ? 'bg-blue-50 border-blue-200' : 'bg-white border-slate-100 hover:border-slate-200'}`} onClick={() => setActivePreviewType('contract')}>
                    <div className="flex items-center justify-between mb-6">
                      <div className="flex items-center gap-3"><div className={`w-12 h-12 rounded-2xl flex items-center justify-center ${contractBuffer ? 'bg-green-100 text-green-600' : 'bg-blue-100 text-blue-600'}`}><FileDown size={24}/></div><div><h3 className="font-bold text-sm">Mẫu Hợp đồng gốc</h3><p className="text-[10px] font-bold text-slate-400 uppercase">{contractBuffer ? contractFileName : 'Chưa thiết lập'}</p></div></div>
                      <button onClick={(e) => { e.stopPropagation(); handleDownloadSample('contract'); }} className="p-2 text-slate-400 hover:text-blue-600 transition-colors"><Download size={18}/></button>
                    </div>
                    <div className="relative group"><input type="file" className="absolute inset-0 opacity-0 cursor-pointer z-10" accept=".docx" onChange={async (e) => { const f = e.target.files?.[0]; if (f) { setContractBuffer(await f.arrayBuffer()); setContractFileName(f.name); setActivePreviewType('contract'); } }} /><div className="py-4 bg-white border-2 border-dashed border-slate-200 rounded-2xl text-center group-hover:border-blue-400 transition-all"><span className="text-xs font-bold text-slate-400 uppercase tracking-widest">CHỌN FILE HĐ (.DOCX)</span></div></div>
                  </div>
                  <div className={`p-6 rounded-[32px] border-2 transition-all cursor-pointer ${activePreviewType === 'request' ? 'bg-slate-900 border-slate-900' : 'bg-white border-slate-100 hover:border-slate-200'}`} onClick={() => setActivePreviewType('request')}>
                    <div className="flex items-center justify-between mb-6">
                      <div className="flex items-center gap-3"><div className={`w-12 h-12 rounded-2xl flex items-center justify-center ${requestBuffer ? 'bg-green-500 text-white' : 'bg-slate-100 text-slate-600'}`}><FileText size={24}/></div><div><h3 className={`font-bold text-sm ${activePreviewType === 'request' ? 'text-white' : 'text-slate-900'}`}>Mẫu Phiếu yêu cầu</h3><p className={`text-[10px] font-bold uppercase text-slate-400`}>{requestBuffer ? requestFileName : 'Chưa thiết lập'}</p></div></div>
                      <button onClick={(e) => { e.stopPropagation(); handleDownloadSample('request'); }} className={`p-2 transition-colors ${activePreviewType === 'request' ? 'text-slate-400 hover:text-white' : 'text-slate-400 hover:text-slate-900'}`}><Download size={18}/></button>
                    </div>
                    <div className="relative group"><input type="file" className="absolute inset-0 opacity-0 cursor-pointer z-10" accept=".docx" onChange={async (e) => { const f = e.target.files?.[0]; if (f) { setRequestBuffer(await f.arrayBuffer()); setRequestFileName(f.name); setActivePreviewType('request'); } }} /><div className={`py-4 border-2 border-dashed rounded-2xl text-center transition-all ${activePreviewType === 'request' ? 'bg-slate-800 border-slate-700 group-hover:border-slate-500' : 'bg-white border-slate-200 group-hover:border-slate-400'}`}><span className={`text-xs font-bold uppercase tracking-widest text-slate-400`}>CHỌN FILE PYC (.DOCX)</span></div></div>
                  </div>
                  <button onClick={() => setStep(AppStep.DASHBOARD)} className="w-full py-5 bg-blue-600 text-white rounded-[24px] font-black shadow-xl flex items-center justify-center gap-2 uppercase tracking-widest hover:scale-[1.02] transition-transform">XÁC NHẬN CẤU HÌNH <ArrowRight size={20}/></button>
                </div>
                <div className="bg-slate-50 rounded-[40px] p-10 h-[650px] overflow-y-auto shadow-inner border border-slate-200 preview-container relative">
                   <div className="absolute top-6 right-8 bg-white/80 backdrop-blur px-4 py-2 rounded-full border border-slate-100 shadow-sm flex items-center gap-2"><Eye size={14} className="text-blue-500"/><span className="text-[10px] font-black uppercase tracking-widest text-slate-500">Xem trước: {activePreviewType === 'contract' ? 'Hợp đồng' : 'Phiếu yêu cầu'}</span></div>
                   {previewHtml ? <div className="document-content pt-12" dangerouslySetInnerHTML={{ __html: previewHtml }} /> : <div className="h-full flex flex-col items-center justify-center text-slate-300"><Files size={48} className="mb-4 opacity-20"/><p className="text-xs font-bold uppercase tracking-widest opacity-40">Chưa có nội dung</p></div>}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {loading && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center">
          <div className="bg-white p-12 rounded-[48px] shadow-2xl text-center max-w-sm animate-in fade-in zoom-in">
            <div className="w-16 h-16 border-[6px] border-blue-600 border-t-transparent rounded-full animate-spin mx-auto mb-6"></div>
            <p className="font-black text-xl uppercase tracking-tighter">Đang xử lý hồ sơ</p>
            <p className="text-slate-400 text-sm font-medium mt-2">{loadingMessage}</p>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
