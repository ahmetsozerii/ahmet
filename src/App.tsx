/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { GoogleGenAI, Type } from "@google/genai";
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { 
  Upload, 
  FileText, 
  Download, 
  Loader2, 
  CheckCircle2, 
  AlertCircle,
  Table as TableIcon,
  Trash2,
  Plus
} from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';

// Initialize PDF.js worker with a more robust approach
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

interface UploadedFile {
  id: string;
  file: File;
  totalPages: number;
  selectedPages: number[];
  pageThumbnails: Record<number, string>;
}

interface ExtractedData {
  id: string;
  atasmanNo: number;
  tarih: string;
  cins: string;
  mahalle: string;
  ilce: string;
  sokak: string;
  kapiNo: string;
  aciklama: string;
  izgaraAdeti: string;
  fullAdres: string;
}

const GEMINI_MODEL = "gemini-3-flash-preview";

export default function App() {
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [data, setData] = useState<ExtractedData[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [processingStats, setProcessingStats] = useState<{ totalTime: number, avgTimePerPage: number } | null>(null);
  const [activeResultPage, setActiveResultPage] = useState<number>(1);
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    document.title = "IZSU PDF&EXCEL CREATION";
  }, []);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFiles(e.dataTransfer.files);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFiles(e.target.files);
    }
  };

  const handleFiles = async (files: FileList | File[]) => {
    const newUploadedFiles: UploadedFile[] = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
        try {
          const arrayBuffer = await file.arrayBuffer();
          const loadingTask = pdfjsLib.getDocument({
            data: arrayBuffer,
            useWorkerFetch: true,
            isEvalSupported: false,
          });
          const pdf = await loadingTask.promise;
          const totalPages = pdf.numPages;
          const selectedPages = Array.from({ length: totalPages }, (_, i) => i + 1);
          
          const fileId = Math.random().toString(36).substring(7);
          
          newUploadedFiles.push({
            id: fileId,
            file,
            totalPages,
            selectedPages,
            pageThumbnails: {}
          });
          
          // Generate thumbnails asynchronously
          (async () => {
            for (let p = 1; p <= totalPages; p++) {
              try {
                const page = await pdf.getPage(p);
                const viewport = page.getViewport({ scale: 0.2 });
                const canvas = document.createElement('canvas');
                const context = canvas.getContext('2d');
                if (context) {
                  canvas.height = viewport.height;
                  canvas.width = viewport.width;
                  await page.render({ canvasContext: context, viewport, canvas }).promise;
                  const thumbUrl = canvas.toDataURL('image/jpeg', 0.5);
                  
                  setUploadedFiles(prev => prev.map(f => {
                    if (f.id === fileId) {
                      return {
                        ...f,
                        pageThumbnails: { ...f.pageThumbnails, [p]: thumbUrl }
                      };
                    }
                    return f;
                  }));
                }
              } catch (e) {
                console.error("Thumbnail error", e);
              }
            }
          })();

        } catch (err) {
          console.error("Error loading PDF to get page count:", err);
          setError(`"${file.name}" okunamadı veya bozuk.`);
        }
      } else {
        setError("Lütfen sadece PDF dosyası yükleyin.");
      }
    }
    setUploadedFiles(prev => [...prev, ...newUploadedFiles]);
    if (newUploadedFiles.length > 0) setError(null);
  };

  const removeFile = (id: string) => {
    setUploadedFiles(prev => prev.filter(f => f.id !== id));
  };

  const togglePageSelection = (fileId: string, pageNum: number) => {
    setUploadedFiles(prev => prev.map(f => {
      if (f.id === fileId) {
        const isSelected = f.selectedPages.includes(pageNum);
        const newSelected = isSelected 
          ? f.selectedPages.filter(p => p !== pageNum)
          : [...f.selectedPages, pageNum].sort((a, b) => a - b);
        return { ...f, selectedPages: newSelected };
      }
      return f;
    }));
  };

  const convertPdfPageToImage = async (pdf: any, pageNum: number): Promise<string> => {
    try {
      const page = await pdf.getPage(pageNum);
      const viewport = page.getViewport({ scale: 1.5 }); // Reduced scale for better performance
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      if (!context) throw new Error("Canvas context could not be created.");
      
      canvas.height = viewport.height;
      canvas.width = viewport.width;

      await page.render({ canvasContext: context, viewport }).promise;
      return canvas.toDataURL('image/jpeg', 0.7).split(',')[1]; // Lower quality for faster processing
    } catch (err: any) {
      console.error(`Sayfa ${pageNum} dönüştürme hatası:`, err);
      throw new Error(`Sayfa ${pageNum} görüntüye dönüştürülemedi.`);
    }
  };

  const processFiles = async () => {
    if (uploadedFiles.length === 0) return;
    
    setIsProcessing(true);
    setError(null);
    setData([]);
    setProcessingStats(null);
    
    const startTime = Date.now();
    let totalProcessedPages = 0;

    try {
      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
        throw new Error("API anahtarı bulunamadı. Lütfen Secrets panelinden GEMINI_API_KEY anahtarını ekleyin.");
      }

      const ai = new GoogleGenAI({ apiKey });
      const allExtractedData: ExtractedData[] = [];
      let currentAtasmanNo = 1;

      const totalPagesToProcess = uploadedFiles.reduce((acc, f) => acc + f.selectedPages.length, 0);
      setProgress({ current: 0, total: totalPagesToProcess });

      for (const uploadedFile of uploadedFiles) {
        if (uploadedFile.selectedPages.length === 0) continue;

        const arrayBuffer = await uploadedFile.file.arrayBuffer();
        const loadingTask = pdfjsLib.getDocument({
          data: arrayBuffer,
          useWorkerFetch: true,
          isEvalSupported: false,
        });
        const pdf = await loadingTask.promise;

        for (const pageNum of uploadedFile.selectedPages) {
          totalProcessedPages++;
          setProgress(prev => ({ ...prev, current: totalProcessedPages }));
          
          const base64Image = await convertPdfPageToImage(pdf, pageNum);

          const prompt = `
            Bu bir el yazısı form sayfasıdır. Sayfada iki ana sütun (Sol ve Sağ) halinde veri listesi bulunmaktadır. 
            Lütfen bu sayfadaki tüm satırları sırasıyla ayıkla ve JSON formatında döndür.
            
            KURALLAR:
            1. Sayfanın üst kısmında MAHALLE ve TARİH bilgileri yer alır. (Örn: YEŞİL MAHALLE, 04.02.2026)
            2. Tablodaki her satır için şu bilgileri bul:
               - TARİH: Sayfanın üstünde yazan tarih.
               - ADRES (YOL İSMİ): "ADRES (SK. CD. BLV.)" sütunundaki yazı.
                 * "sk", "sk.", "sok", "sok." görüyorsan türü "SOKAK" olarak belirle.
                 * "cd", "cd.", "cad", "cad." görüyorsan türü "CADDESİ" olarak belirle.
                 * "blv", "blv." görüyorsan türü "BULVARI" olarak belirle.
                 * Yol isminde "YANI" kelimesi geçiyorsa, bu kelimeyi AÇIKLAMA kısmına taşı, adresten sil.
                 * Eğer "-" işaretinden sonra bir şey yazıyorsa, bu kısmı AÇIKLAMA olarak ayır.
                 * Eğer adreste "650 sk. 658" veya "ATATÜRK CAD. 56" gibi fazladan bir sayı varsa, bu son sayıyı (658 veya 56) AÇIKLAMA kısmına taşı ve adresten sil.
                 * Eğer adreste "KUS", "KURS", "kvs", "KVS" gibi kelimeler varsa, bunları "KVŞ" olarak düzelt.
               - KAPI NO: "KAPI NO" sütunundaki değer. 
                 * Eğer kapı numarasında hem sayı hem de "KVŞ" (veya KUS, KVS) varsa, sayıyı AÇIKLAMA kısmına taşı ve KAPI NO kısmına sadece "KVŞ" yaz.
                 * Eğer kapı numarasında "ÖNÜ", "ARKASI", "YANI", "KARŞISI" gibi kelimeler varsa, bu kelimeleri AÇIKLAMA kısmına taşı ve kapı numarasından sil.
               - IZGARA ADETİ: "IZGARA ADETİ" sütunundaki sayısal miktar.
            3. " " (den den) işareti varsa, bu bir üstteki satırla aynı veri demektir. Lütfen bu veriyi bir üsttekinden kopyalayarak doldur. Kesinlikle çıktı JSON'da " işareti bırakma, gerçek değeri yaz.
            4. Tüm metinleri BÜYÜK HARF yap.
            5. Mahalle isminde "MH" veya "MAHALLE" geçiyorsa sadece ismini al (Örn: "YEŞİL"). İlçe ismi genellikle mahalle isminin yanında veya altında parantez içinde veya ayrı yazılır (Örn: "KARABAĞLAR").
            
            Döndürmen gereken JSON yapısı (ARRAY):
            [{
              "tarih": "GG.AA.YYYY",
              "mahalle": "...",
              "ilce": "...",
              "yolIsmi": "...",
              "yolTuru": "BULVARI | CADDESİ | SOKAK",
              "kapiNo": "...",
              "aciklama": "...",
              "izgaraAdeti": "..."
            }]
          `;

          const response = await ai.models.generateContent({
            model: GEMINI_MODEL,
            contents: [
              {
                parts: [
                  { text: prompt },
                  { inlineData: { mimeType: "image/jpeg", data: base64Image } }
                ]
              }
            ],
            config: {
              responseMimeType: "application/json",
              responseSchema: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    tarih: { type: Type.STRING },
                    mahalle: { type: Type.STRING },
                    ilce: { type: Type.STRING },
                    yolIsmi: { type: Type.STRING },
                    yolTuru: { type: Type.STRING },
                    kapiNo: { type: Type.STRING },
                    aciklama: { type: Type.STRING },
                    izgaraAdeti: { type: Type.STRING }
                  },
                  required: ["tarih", "mahalle", "ilce", "yolIsmi", "yolTuru", "kapiNo", "izgaraAdeti"]
                }
              }
            }
          });

          const pageData = JSON.parse(response.text || "[]");
          
          pageData.forEach((item: any) => {
            const mahalle = item.mahalle?.toUpperCase().replace(/"/g, '') || "";
            const ilce = item.ilce?.toUpperCase().replace(/"/g, '') || "";
            let yolIsmi = item.yolIsmi?.toUpperCase().replace(/"/g, '') || "";
            const yolTuru = item.yolTuru?.toUpperCase().replace(/"/g, '') || "SOKAK";
            let rawKapiNo = item.kapiNo?.toUpperCase().replace(/"/g, '') || "";
            let aciklama = item.aciklama?.toUpperCase().replace(/"/g, '') || "";

            // Fix OCR misread "KUS" to "KVŞ" (Kavşak)
            const kvsRegex = /\b(KUS|KURS|KVS|kvs|kvs\.|kvs,|KVS\.|KVS,|KUS\.|KUS,|KURS\.|KURS,|KVŞ|KVŞ\.|KVŞ,)\b/gi;
            
            // Move KVŞ to KAPI NO and remove from AÇIKLAMA and YOL İSMİ
            let hasKvs = false;
            if (kvsRegex.test(rawKapiNo) || kvsRegex.test(yolIsmi) || kvsRegex.test(aciklama)) {
              hasKvs = true;
              rawKapiNo = rawKapiNo.replace(kvsRegex, '').trim();
              yolIsmi = yolIsmi.replace(kvsRegex, '').trim();
              aciklama = aciklama.replace(kvsRegex, '').trim();
            }

            // Extract YANI first so it doesn't interfere with finding the last number
            let cleanYolIsmi = yolIsmi.split("YANI")[0].trim();
            const extraFromYol = yolIsmi.includes("YANI") ? "YANI " + yolIsmi.split("YANI")[1].trim() : "";

            // Handle extra number in address (e.g., "650 sk. 658" or "ATATÜRK 56")
            const lastNumberMatch = cleanYolIsmi.match(/\s+(\d+)\s*$/);
            if (lastNumberMatch) {
              const extraNumber = lastNumberMatch[1];
              aciklama = (aciklama + " " + extraNumber).trim();
              cleanYolIsmi = cleanYolIsmi.replace(new RegExp(`\\s+${extraNumber}\\s*$`), '').trim();
            } else {
              const addressNumbers = cleanYolIsmi.match(/\b\d+\b/g);
              if (addressNumbers && addressNumbers.length > 1 && /^\s*\d+/.test(cleanYolIsmi)) {
                const secondNumber = addressNumbers[1];
                if (!cleanYolIsmi.includes(`${secondNumber}.`)) {
                  aciklama = (aciklama + " " + secondNumber).trim();
                  cleanYolIsmi = cleanYolIsmi.replace(new RegExp(`\\b${secondNumber}\\b`), '').trim();
                }
              }
            }

            // Handle KVŞ priority in KAPI NO
            if (hasKvs) {
               if (rawKapiNo.trim()) {
                 aciklama = (aciklama + " " + rawKapiNo.trim()).trim();
               }
               rawKapiNo = "KVŞ";
            }

            // Handle terms in KAPI NO
            const kapiNoTerms = ["ÖNÜ", "ARKASI", "YANI", "KARŞISI"];
            let cleanKapiNo = rawKapiNo;
            let extraFromKapi = "";
            
            kapiNoTerms.forEach(term => {
              if (cleanKapiNo.includes(term)) {
                extraFromKapi += " " + term;
                cleanKapiNo = cleanKapiNo.replace(term, "").trim();
              }
            });
            
            let finalAciklama = (extraFromYol + " " + extraFromKapi + " " + aciklama).trim();
            
            // Ensure KVŞ is completely removed from the final aciklama string
            finalAciklama = finalAciklama.replace(/\b(KUS|KURS|KVS|kvs|kvs\.|kvs,|KVS\.|KVS,|KUS\.|KUS,|KURS\.|KURS,|KVŞ|KVŞ\.|KVŞ,)\b/gi, '').trim();
            
            // Append " SOK. KESİŞİM" if description is not empty and doesn't contain specific terms
            if (finalAciklama) {
              const hasExclusion = ["YANI", "KARŞISI", "ÖNÜ", "ARKASI"].some(term => finalAciklama.includes(term));
              if (!hasExclusion) {
                finalAciklama += " SOK. KESİŞİM";
              }
            }

            const fullAdres = `${mahalle} MAH. ${cleanYolIsmi} ${yolTuru} NO: ${cleanKapiNo} ${ilce}`;

            allExtractedData.push({
              id: Math.random().toString(36).substring(7),
              atasmanNo: currentAtasmanNo,
              tarih: item.tarih?.toUpperCase().replace(/"/g, '') || "",
              cins: "TEKLİ",
              mahalle,
              ilce,
              sokak: cleanYolIsmi,
              kapiNo: cleanKapiNo,
              aciklama: finalAciklama,
              izgaraAdeti: item.izgaraAdeti || "",
              fullAdres
            });
          });
          
          currentAtasmanNo++; // Increment per page
        }
      }

      setData(allExtractedData);
      
      const endTime = Date.now();
      const durationMs = endTime - startTime;
      setProcessingStats({
        totalTime: durationMs,
        avgTimePerPage: totalProcessedPages > 0 ? durationMs / totalProcessedPages : 0
      });
      
      if (allExtractedData.length > 0) {
        setActiveResultPage(allExtractedData[0].atasmanNo);
      }

    } catch (err: any) {

      console.error("PDF İşleme Hatası:", err);
      let message = "Dosya işlenirken bir hata oluştu.";
      if (err.message?.includes("API key")) {
        message = "API anahtarı geçersiz veya eksik.";
      } else if (err.message?.includes("PDF")) {
        message = "PDF dosyası okunamadı veya bozuk.";
      } else if (err.message) {
        message = `Hata: ${err.message}`;
      }
      setError(message);
    } finally {
      setIsProcessing(false);
    }
  };

  const exportToExcel = () => {
    const worksheetData = data.map(item => ({
      'ATAŞMAN NO': item.atasmanNo,
      'TARİH': item.tarih,
      'CİNS': item.cins,
      'ADRES': item.fullAdres,
      'AÇIKLAMA': item.aciklama,
      'IZGARA ADETİ': item.izgaraAdeti
    }));

    const worksheet = XLSX.utils.json_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Veriler");
    
    // Generate filename IZSUVERI1, IZSUVERI2, etc. based on localStorage counter
    let fileCounter = parseInt(localStorage.getItem('izsuveri_counter') || '1', 10);
    const fileName = `IZSUVERI${fileCounter}.xlsx`;
    localStorage.setItem('izsuveri_counter', (fileCounter + 1).toString());
    
    XLSX.writeFile(workbook, fileName);
  };

  const handleCellChange = (id: string, field: keyof ExtractedData, value: string) => {
    setData(prev => prev.map(item => {
      if (item.id === id) {
        if (field === 'atasmanNo') {
          return { ...item, [field]: parseInt(value) || 0 };
        }
        return { ...item, [field]: value };
      }
      return item;
    }));
  };

  const uniqueAtasmanNos = Array.from(new Set(data.map(d => Number(d.atasmanNo)))).sort((a: number, b: number) => a - b);
  const currentData = data.filter(d => Number(d.atasmanNo) === activeResultPage);

  return (
    <div className="min-h-screen bg-[#0F172A] text-slate-100 font-sans selection:bg-emerald-500 selection:text-white p-4 md:p-8 relative overflow-hidden">
      {/* Background Decorative Elements */}
      <div className="absolute top-0 left-0 w-full h-full overflow-hidden pointer-events-none z-0">
        <div className="absolute -top-[10%] -left-[10%] w-[40%] h-[40%] bg-emerald-500/10 blur-[120px] rounded-full" />
        <div className="absolute top-[20%] -right-[5%] w-[30%] h-[30%] bg-indigo-500/10 blur-[100px] rounded-full" />
        <div className="absolute -bottom-[10%] left-[20%] w-[50%] h-[50%] bg-amber-500/5 blur-[150px] rounded-full" />
      </div>

      <div className="max-w-6xl mx-auto relative z-10">
        {/* Header */}
        <header className="mb-12 flex flex-col md:flex-row md:items-end justify-between gap-6 border-b border-slate-700 pb-8">
          <motion.div
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
          >
            <div className="flex items-center gap-3 mb-4">
              <div className="w-12 h-12 bg-gradient-to-br from-emerald-400 to-emerald-600 rounded-2xl flex items-center justify-center shadow-lg shadow-emerald-500/20">
                <FileText className="text-white" size={24} />
              </div>
              <h1 className="text-4xl md:text-5xl font-serif font-bold tracking-tight bg-clip-text text-transparent bg-gradient-to-r from-white to-slate-400">
                IZSU PDF&EXCEL <span className="italic font-normal text-emerald-400">CREATION</span>
              </h1>
            </div>
            <p className="text-xs uppercase tracking-[0.3em] text-emerald-500/80 font-bold">Yapay Zeka Destekli Veri Otomasyonu</p>
          </motion.div>
          
          <motion.div 
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            className="flex items-center gap-4"
          >
            <div className="text-right hidden md:block">
              <p className="text-[10px] uppercase tracking-widest text-slate-500">Sistem Durumu</p>
              <div className="flex items-center justify-end gap-2">
                <div className={`w-2 h-2 rounded-full ${isProcessing ? 'bg-amber-500 animate-pulse' : 'bg-emerald-500'}`} />
                <p className="text-xs font-mono text-slate-300">{isProcessing ? 'İŞLENİYOR' : 'AKTİF'}</p>
              </div>
            </div>
          </motion.div>
        </header>

        {/* Upload Section */}
        <section className="mb-12">
          <motion.div 
            whileHover={{ scale: 1.01 }}
            whileTap={{ scale: 0.99 }}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            className={`
              relative group cursor-pointer
              border-2 border-dashed rounded-[2.5rem] p-12
              transition-all duration-500 ease-out
              ${isDragging ? 'bg-slate-800/60 border-emerald-400 scale-[1.02]' : ''}
              ${uploadedFiles.length > 0 
                ? 'bg-slate-800/50 border-emerald-500/50 shadow-2xl shadow-emerald-500/10' 
                : 'bg-slate-800/20 border-slate-700 hover:border-emerald-500/30 hover:bg-slate-800/30'}
            `}
            onClick={() => fileInputRef.current?.click()}
          >
            <input 
              type="file" 
              ref={fileInputRef}
              className="hidden" 
              accept=".pdf,application/pdf"
              multiple
              onChange={handleFileChange}
            />
            
            <div className="flex flex-col items-center text-center gap-6">
              <div className={`
                w-20 h-20 rounded-3xl flex items-center justify-center
                transition-all duration-500 group-hover:rotate-6
                ${uploadedFiles.length > 0 
                  ? 'bg-emerald-500 text-white shadow-xl shadow-emerald-500/30' 
                  : 'bg-slate-700 text-slate-400'}
              `}>
                {uploadedFiles.length > 0 ? <CheckCircle2 size={40} /> : <Upload size={40} />}
              </div>
              
              <div>
                <h3 className="text-2xl font-semibold mb-2 text-white">
                  {uploadedFiles.length > 0 ? `${uploadedFiles.length} Dosya Seçildi` : 'PDF Dosyalarını Buraya Sürükleyin'}
                </h3>
                <p className="text-sm text-slate-400 max-w-xs mx-auto">
                  {uploadedFiles.length > 0 
                    ? 'Aşağıdan sayfaları seçip analizi başlatabilirsiniz.' 
                    : 'Veya bilgisayarınızdan dosya seçmek için bu alana tıklayın.'}
                </p>
              </div>

              {uploadedFiles.length > 0 && !isProcessing && (
                <motion.button
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  onClick={(e) => {
                    e.stopPropagation();
                    processFiles();
                  }}
                  className="mt-4 px-10 py-4 bg-emerald-500 text-white rounded-2xl text-sm font-bold tracking-wider hover:bg-emerald-400 transition-all flex items-center gap-3 shadow-lg shadow-emerald-500/20 active:scale-95"
                >
                  <FileText size={20} />
                  ANALİZİ BAŞLAT
                </motion.button>
              )}
            </div>
          </motion.div>

          {/* Uploaded Files List */}
          {uploadedFiles.length > 0 && (
            <div className="mt-8 space-y-4">
              {uploadedFiles.map(file => (
                <div key={file.id} className="bg-slate-800/50 rounded-2xl p-4 border border-slate-700 flex flex-col md:flex-row gap-4 items-start md:items-center justify-between">
                  <div className="flex items-center gap-3">
                    <FileText className="text-emerald-400" size={24} />
                    <div>
                      <p className="text-sm font-medium text-slate-200">{file.file.name}</p>
                      <p className="text-xs text-slate-500">{(file.file.size / 1024 / 1024).toFixed(2)} MB • {file.totalPages} Sayfa</p>
                    </div>
                  </div>
                  
                  {!isProcessing && data.length === 0 && (
                    <div className="flex-1 w-full md:w-auto overflow-x-auto">
                      <div className="flex gap-4 pb-4">
                        {Array.from({ length: file.totalPages }, (_, i) => i + 1).map(pageNum => {
                          const isSelected = file.selectedPages.includes(pageNum);
                          const thumb = file.pageThumbnails[pageNum];
                          return (
                            <div key={pageNum} className="flex flex-col items-center gap-2">
                              <button
                                onClick={() => togglePageSelection(file.id, pageNum)}
                                className={`min-w-[40px] h-10 rounded-lg text-xs font-bold transition-colors ${
                                  isSelected 
                                    ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' 
                                    : 'bg-slate-900/50 text-slate-500 border border-slate-700 hover:bg-slate-700'
                                }`}
                              >
                                {pageNum}
                              </button>
                              {thumb ? (
                                <img 
                                  src={thumb} 
                                  alt={`Sayfa ${pageNum}`} 
                                  className={`w-16 h-20 object-cover rounded border ${isSelected ? 'border-emerald-500/50' : 'border-slate-700 opacity-50'} transition-all cursor-pointer`}
                                  onClick={() => togglePageSelection(file.id, pageNum)}
                                  referrerPolicy="no-referrer"
                                />
                              ) : (
                                <div className={`w-16 h-20 rounded border border-slate-700 bg-slate-800/50 flex items-center justify-center ${isSelected ? '' : 'opacity-50'}`}>
                                  <Loader2 className="animate-spin text-slate-500" size={14} />
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {!isProcessing && data.length === 0 && (
                    <button 
                      onClick={() => removeFile(file.id)}
                      className="p-2 text-red-400 hover:bg-red-400/10 rounded-lg transition-colors"
                      title="Dosyayı İptal Et"
                    >
                      <Trash2 size={20} />
                    </button>
                  )}
                  
                  {data.length > 0 && (
                    <div className="flex items-center gap-4">
                      <div className="px-3 py-1 bg-emerald-500/10 text-emerald-400 rounded-lg text-xs font-bold border border-emerald-500/20">
                        {file.selectedPages.length} Sayfa İşlendi
                      </div>
                      <button 
                        onClick={() => removeFile(file.id)}
                        className="p-2 text-red-400 hover:bg-red-400/10 rounded-lg transition-colors"
                        title="Dosyayı İptal Et"
                      >
                        <Trash2 size={20} />
                      </button>
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}


          {isProcessing && (
            <div className="mt-12 space-y-6 max-w-2xl mx-auto">
              <div className="flex items-center justify-between text-[10px] uppercase tracking-[0.2em] text-emerald-400 font-bold">
                <span>İşlem Durumu</span>
                <span>Sayfa {progress.current} / {progress.total}</span>
              </div>
              <div className="h-2 w-full bg-slate-800 rounded-full overflow-hidden border border-slate-700">
                <motion.div 
                  className="h-full bg-gradient-to-r from-emerald-500 to-indigo-500"
                  initial={{ width: 0 }}
                  animate={{ width: `${(progress.current / progress.total) * 100}%` }}
                  transition={{ type: "spring", bounce: 0, duration: 0.5 }}
                />
              </div>
              <div className="flex flex-col items-center gap-3">
                <div className="flex items-center gap-3 text-sm font-medium text-slate-300">
                  <Loader2 className="animate-spin text-emerald-500" size={18} />
                  <span>Gemini AI belgeleri derinlemesine analiz ediyor...</span>
                </div>
                <p className="text-[10px] text-slate-500 uppercase tracking-widest">Lütfen pencereyi kapatmayın</p>
              </div>
            </div>
          )}

          {error && (
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              className="mt-8 p-6 bg-red-500/10 border border-red-500/20 rounded-3xl flex items-center gap-4 text-red-400"
            >
              <div className="w-12 h-12 rounded-2xl bg-red-500/20 flex items-center justify-center shrink-0">
                <AlertCircle size={24} />
              </div>
              <div>
                <p className="text-sm font-bold uppercase tracking-wider mb-1">Hata Oluştu</p>
                <p className="text-sm opacity-80">{error}</p>
              </div>
            </motion.div>
          )}
        </section>

        {/* Results Section */}
        <AnimatePresence>
          {data.length > 0 && (
            <motion.section
              initial={{ opacity: 0, y: 40 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 40 }}
              className="space-y-8"
            >
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                <div className="flex items-center gap-4">
                  <div className="w-14 h-14 rounded-2xl bg-indigo-500/20 text-indigo-400 flex items-center justify-center shadow-inner">
                    <TableIcon size={28} />
                  </div>
                  <div>
                    <h2 className="text-2xl font-bold text-white">Analiz Sonuçları</h2>
                    <p className="text-xs text-emerald-500 font-bold uppercase tracking-widest">{data.length} Kayıt Başarıyla Ayıklandı</p>
                  </div>
                </div>
                
                {processingStats && (
                  <div className="hidden lg:flex items-center gap-4 px-4 py-2 bg-slate-800/50 rounded-xl border border-slate-700">
                    <div className="text-right">
                      <p className="text-[10px] text-slate-500 uppercase tracking-wider">Toplam Süre</p>
                      <p className="text-sm font-mono text-emerald-400">{(processingStats.totalTime / 1000).toFixed(1)}s</p>
                    </div>
                    <div className="w-px h-8 bg-slate-700" />
                    <div className="text-left">
                      <p className="text-[10px] text-slate-500 uppercase tracking-wider">Sayfa Başına</p>
                      <p className="text-sm font-mono text-emerald-400">{(processingStats.avgTimePerPage / 1000).toFixed(1)}s</p>
                    </div>
                  </div>
                )}

                <div className="flex items-center gap-4 px-6 py-3 bg-emerald-500/10 rounded-2xl border border-emerald-500/20 shadow-inner">
                  <div className="text-center">
                    <p className="text-[10px] text-emerald-500 uppercase tracking-widest font-bold mb-1">Genel Toplam Izgara</p>
                    <p className="text-2xl font-black text-emerald-400 leading-none">
                      {data.reduce((sum, item) => sum + (parseInt(item.izgaraAdeti) || 0), 0)}
                    </p>
                  </div>
                </div>
                
                <div className="flex items-center gap-3">
                  <button
                    onClick={() => setData([])}
                    className="p-3 bg-slate-800 text-slate-400 rounded-2xl hover:bg-red-500/10 hover:text-red-400 transition-all border border-slate-700"
                    title="Listeyi Temizle"
                  >
                    <Trash2 size={20} />
                  </button>
                  <button
                    onClick={exportToExcel}
                    className="px-8 py-4 bg-amber-500 text-slate-900 rounded-2xl text-sm font-black uppercase tracking-wider flex items-center gap-3 hover:bg-amber-400 transition-all shadow-xl shadow-amber-500/20 active:scale-95"
                  >
                    <Download size={20} />
                    EXCEL DOSYASINI İNDİR
                  </button>
                </div>
              </div>

              {uniqueAtasmanNos.length > 1 && (
                <div className="flex gap-2 pb-2 overflow-x-auto">
                  {uniqueAtasmanNos.map(no => {
                    const rowCount = data.filter(d => Number(d.atasmanNo) === no).length;
                    return (
                      <button
                        key={no}
                        onClick={() => setActiveResultPage(no)}
                        className={`px-6 py-2 rounded-xl transition-all whitespace-nowrap flex flex-col items-center ${
                          activeResultPage === no
                            ? 'bg-emerald-500 text-white shadow-lg shadow-emerald-500/20'
                            : 'bg-slate-800 text-slate-400 hover:bg-slate-700 border border-slate-700'
                        }`}
                      >
                        <span className="text-sm font-bold">Sayfa {no}</span>
                        <span className={`text-[10px] font-medium ${activeResultPage === no ? 'text-emerald-100' : 'text-slate-500'}`}>
                          {rowCount} Kayıt
                        </span>
                      </button>
                    );
                  })}
                </div>
              )}

              <div className="bg-slate-800/40 rounded-[2rem] shadow-2xl overflow-hidden border border-slate-700 backdrop-blur-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-900/50 border-b border-slate-700">
                        <th className="p-5 text-[10px] uppercase tracking-[0.2em] font-black text-slate-500">No</th>
                        <th className="p-5 text-[10px] uppercase tracking-[0.2em] font-black text-slate-500">Tarih</th>
                        <th className="p-5 text-[10px] uppercase tracking-[0.2em] font-black text-slate-500">Cins</th>
                        <th className="p-5 text-[10px] uppercase tracking-[0.2em] font-black text-slate-500">Adres Bilgisi</th>
                        <th className="p-5 text-[10px] uppercase tracking-[0.2em] font-black text-slate-500">Açıklama</th>
                        <th className="p-5 text-[10px] uppercase tracking-[0.2em] font-black text-slate-500">
                          <div className="flex flex-col items-center">
                            <span>Adet</span>
                            <span className="text-emerald-400 mt-1 bg-emerald-500/10 px-2 py-0.5 rounded-md border border-emerald-500/20">
                              Toplam: {currentData.reduce((sum, item) => sum + (parseInt(item.izgaraAdeti) || 0), 0)}
                            </span>
                          </div>
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-700/50">
                      {currentData.map((item, idx) => (
                        <motion.tr 
                          key={item.id}
                          initial={{ opacity: 0, x: -10 }}
                          animate={{ opacity: 1, x: 0 }}
                          transition={{ delay: idx * 0.03 }}
                          className="hover:bg-white/5 transition-colors group"
                        >
                          <td className="p-5 font-mono text-xs text-slate-500 group-hover:text-emerald-400 transition-colors">
                            <input 
                              type="number"
                              value={item.atasmanNo}
                              onChange={(e) => handleCellChange(item.id, 'atasmanNo', e.target.value)}
                              className="bg-transparent border-none focus:ring-1 focus:ring-emerald-500/30 rounded w-12 p-1 text-center"
                            />
                          </td>
                          <td className="p-5 text-sm font-medium">
                            <input 
                              type="text"
                              value={item.tarih}
                              onChange={(e) => handleCellChange(item.id, 'tarih', e.target.value)}
                              className="bg-transparent border-none focus:ring-1 focus:ring-emerald-500/30 rounded w-24 p-1"
                            />
                          </td>
                          <td className="p-5">
                            <input 
                              type="text"
                              value={item.cins}
                              onChange={(e) => handleCellChange(item.id, 'cins', e.target.value)}
                              className="bg-transparent border-none focus:ring-1 focus:ring-emerald-500/30 rounded w-16 p-1 text-indigo-400 font-black uppercase text-[10px]"
                            />
                          </td>
                          <td className="p-5">
                            <textarea 
                              rows={2}
                              value={item.fullAdres}
                              onChange={(e) => handleCellChange(item.id, 'fullAdres', e.target.value)}
                              className="bg-transparent border-none focus:ring-1 focus:ring-emerald-500/30 rounded w-full p-1 text-slate-200 font-medium resize-none"
                            />
                          </td>
                          <td className="p-5">
                            <input 
                              type="text"
                              value={item.aciklama}
                              onChange={(e) => handleCellChange(item.id, 'aciklama', e.target.value)}
                              className="bg-transparent border-none focus:ring-1 focus:ring-emerald-500/30 rounded w-full p-1 text-amber-400 text-xs font-medium"
                              placeholder="Açıklama..."
                            />
                          </td>
                          <td className="p-5">
                            <div className="flex items-center justify-center">
                              <input 
                                type="text"
                                value={item.izgaraAdeti}
                                onChange={(e) => handleCellChange(item.id, 'izgaraAdeti', e.target.value)}
                                className="bg-transparent border-none focus:ring-1 focus:ring-emerald-500/30 rounded w-10 p-1 text-center font-bold text-emerald-400"
                              />
                            </div>
                          </td>
                        </motion.tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

            </motion.section>
          )}
        </AnimatePresence>

        {/* Empty State */}
        {uploadedFiles.length === 0 && !isProcessing && data.length === 0 && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            className="py-32 flex flex-col items-center text-center"
          >
            <div className="relative mb-8">
              <div className="absolute inset-0 bg-emerald-500/20 blur-3xl rounded-full" />
              <div className="relative w-32 h-32 bg-slate-800 border-2 border-dashed border-slate-700 rounded-[2.5rem] flex items-center justify-center text-slate-600">
                <FileText size={48} />
              </div>
            </div>
            <h3 className="text-2xl font-serif italic text-slate-400 mb-2">Henüz bir belge yüklenmedi</h3>
            <p className="text-xs uppercase tracking-[0.3em] text-slate-600 font-bold">Analiz için bir PDF dosyası seçin</p>
          </motion.div>
        )}
      </div>

      {/* Footer Info */}
      <footer className="mt-24 border-t border-slate-800 pt-12 pb-12 text-center">
        <div className="flex items-center justify-center gap-6 mb-6 opacity-40">
          <div className="h-px w-12 bg-slate-700" />
          <p className="text-[10px] uppercase tracking-[0.4em] font-black text-slate-400">
            Enterprise Data Intelligence
          </p>
          <div className="h-px w-12 bg-slate-700" />
        </div>
        <p className="text-[9px] text-slate-600 uppercase tracking-widest">
          Powered by Google Gemini 3 Flash & PDF.js Engine
        </p>
      </footer>
    </div>
  );
}
