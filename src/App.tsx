/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useRef } from 'react';
import { jsPDF } from 'jspdf';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { 
  FileText, 
  Settings, 
  Download, 
  Users, 
  Layout, 
  ChevronRight, 
  ChevronDown,
  Info,
  AlertCircle,
  CheckCircle2,
  Printer,
  Upload,
  Trash2,
  Table as TableIcon,
  Eye,
  Save,
  Plus,
  Bookmark,
  History
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';
import { OpticField, StudentData, Calibration, SavedProfile } from './types';

const DEFAULT_FMT = `53=42=04=D=*= =
20=29=38=43=K=D=0123456789=X2=KURUM KODU=
07=16=26=30=K=D=0123456789=X2=ÖĞRENCİ NO=
01=29=02=11=K=D=ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ=X2=ADI=
01=29=12=21=K=D=ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ=X2=SOYADI=
01=04=23=23=K=D=5678=X2=SINIF=
01=29=24=24=K=D=ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ=X2=ŞUBE=
20=29=26=36=K=D=0123456789=X2=TC/TEL=
36=36=04=08=K=Y=K   E=X2=CİNSİYET=
39=41=07=07=K=D=1 2=X2=OTURUM=
32=32=04=07=K=Y=A  B=X2=KİTAPÇIK TÜRÜ=
34=53=13=16=K=Y=ABCD=X2=TÜRKÇE=
34=53=18=21=K=Y=ABCD=X2=SOSYAL=
34=53=23=26=K=Y=ABCD=X2=DİN=
34=53=28=31=K=Y=ABCD=X2=İNGİLİZCE=
34=53=35=38=K=Y=ABCD=X2=MATEMATİK=
34=53=40=43=K=Y=ABCD=X2=FEN=`;

const DEFAULT_STUDENTS = `GÜLİSTAN	ŞAHİN	7	755676	7	A
ASLI NUR	SALTEK	23	755676	7	A
HATİCE KÜBRA	ALTUNKAYA	77	755676	7	A
NEVRA RANA	KADİRHAN	95	755676	7	A
FATİH	YÜCEL	99	755676	7	A`;

export default function App() {
  const [fmtText, setFmtText] = useState(DEFAULT_FMT);
  const [studentText, setStudentText] = useState(DEFAULT_STUDENTS);
  const [hasHeader, setHasHeader] = useState(false);
  const [fieldMappings, setFieldMappings] = useState<Record<string, number>>({
    'ADI': 0,
    'SOYADI': 1,
    'ÖĞRENCİ NO': 2,
    'KURUM KODU': 3,
    'SINIF': 4,
    'ŞUBE': 5,
  });
  const [calibration, setCalibration] = useState<Calibration>({
    offsetX: 15, // mm
    offsetY: 15, // mm
    bubbleSpacing: 25.4 / 6, // 1/6 inch in mm (~4.233)
    bubbleSize: 3.5, // mm
  });
  const [pageSize, setPageSize] = useState({ width: 21.0, height: 29.7 }); // cm
  const [isGenerating, setIsGenerating] = useState(false);
  const [activeTab, setActiveTab] = useState<'data' | 'format' | 'settings' | 'preview'>('data');
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'error' } | null>(null);
  const [profiles, setProfiles] = useState<SavedProfile[]>(() => {
    const saved = localStorage.getItem('optic_form_profiles');
    return saved ? JSON.parse(saved) : [];
  });
  const [profileName, setProfileName] = useState('');
  const [isAuthorized, setIsAuthorized] = useState(() => {
    return localStorage.getItem('app_authorized') === 'true';
  });
  const [passwordInput, setPasswordInput] = useState('');
  const [passwordError, setPasswordError] = useState(false);

  const handlePasswordSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (passwordInput === 'TayTay+2494') {
      setIsAuthorized(true);
      localStorage.setItem('app_authorized', 'true');
      setPasswordError(false);
    } else {
      setPasswordError(true);
      showToast('Hatalı parola!', 'error');
    }
  };

  const showToast = (message: string, type: 'success' | 'error' = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3000);
  };

  const saveProfile = () => {
    if (!profileName.trim()) return;
    const newProfile: SavedProfile = {
      id: crypto.randomUUID(),
      name: profileName,
      fmtText,
      calibration,
      fieldMappings,
      pageSize,
      createdAt: Date.now(),
    };
    const updated = [...profiles, newProfile];
    setProfiles(updated);
    localStorage.setItem('optic_form_profiles', JSON.stringify(updated));
    setProfileName('');
    showToast('Profil başarıyla kaydedildi');
  };

  const loadProfile = (profile: SavedProfile) => {
    setFmtText(profile.fmtText);
    setCalibration(profile.calibration);
    setFieldMappings(profile.fieldMappings);
    setPageSize(profile.pageSize);
    setActiveTab('settings');
    showToast('Profil yüklendi');
  };

  const deleteProfile = (id: string) => {
    const updated = profiles.filter(p => p.id !== id);
    setProfiles(updated);
    localStorage.setItem('optic_form_profiles', JSON.stringify(updated));
    showToast('Profil silindi');
  };

  const downloadSampleExcel = () => {
    const data = [
      ['ADI', 'SOYADI', 'ÖĞRENCİ NO', 'KURUM KODU', 'SINIF', 'ŞUBE'],
      ['GÜLİSTAN', 'ŞAHİN', '7', '755676', '7', 'A'],
      ['ASLI NUR', 'SALTEK', '23', '755676', '7', 'A'],
      ['HATİCE KÜBRA', 'ALTUNKAYA', '77', '755676', '7', 'A'],
      ['NEVRA RANA', 'KADİRHAN', '95', '755676', '7', 'A'],
      ['FATİH', 'YÜCEL', '99', '755676', '7', 'A']
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Öğrenciler");
    XLSX.writeFile(wb, "ornek_ogrenci_listesi.xlsx");
    showToast('Örnek Excel indirildi');
  };

  const generateTestPage = async () => {
    const doc = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4'
    });

    // Load a font that supports Turkish characters
    try {
      const fontUrl = 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf';
      const response = await fetch(fontUrl);
      const buffer = await response.arrayBuffer();
      const base64 = btoa(new Uint8Array(buffer).reduce((data, byte) => data + String.fromCharCode(byte), ''));
      doc.addFileToVFS('Roboto-Medium.ttf', base64);
      doc.addFont('Roboto-Medium.ttf', 'Roboto', 'normal');
      doc.setFont('Roboto');
    } catch (e) {
      console.warn('Font loading failed, falling back to helvetica', e);
      doc.setFont('helvetica', 'bold');
    }

    const width = 210;
    const height = 297;

    // Draw grid
    doc.setDrawColor(220, 220, 220);
    doc.setLineWidth(0.1);

    // Vertical lines every 10mm
    for (let x = 0; x <= width; x += 10) {
      doc.line(x, 0, x, height);
      if (x > 0 && x < width) {
        doc.setFontSize(6);
        doc.setTextColor(150, 150, 150);
        doc.text(x.toString(), x, 7, { align: 'center' });
      }
    }

    // Horizontal lines every 10mm
    for (let y = 0; y <= height; y += 10) {
      doc.line(0, y, width, y);
      if (y > 0 && y < height) {
        doc.setFontSize(6);
        doc.setTextColor(150, 150, 150);
        doc.text(y.toString(), 4, y + 1);
      }
    }

    // Starting point based on current calibration
    doc.setFillColor(255, 0, 0);
    doc.circle(calibration.offsetX, calibration.offsetY, 1, 'F');
    doc.setFontSize(10);
    doc.setTextColor(255, 0, 0);
    doc.text(`Baslangic (X:${calibration.offsetX}, Y:${calibration.offsetY})`, calibration.offsetX + 3, calibration.offsetY + 1);

    // Instructions
    doc.setFontSize(16);
    doc.setTextColor(0, 0, 0);
    doc.text("OPTIK FORM KALIBRASYON TEST SAYFASI", width / 2, 150, { align: 'center' });
    
    doc.setFontSize(12);
    doc.text("Lutfen yazicidan cikarip orijinal optik formunuzun ustune koyarak isiga tutun.", width / 2, 160, { align: 'center' });
    doc.text("Kırmızı nokta soldaki siyah kılavuz çizgilerin", width / 2, 170, { align: 'center' });
    doc.text("birinci sütün birinci baloncuğa işaretleyin.", width / 2, 178, { align: 'center' });

    doc.save('kalibrasyon_test_sayfasi.pdf');
    showToast('Kalibrasyon test sayfası indirildi');
  };

  const gridConfig = useMemo(() => {
    const firstLine = fmtText.split('\n')[0];
    if (!firstLine || !firstLine.includes('=')) return { totalRows: 53, totalCols: 42 };
    const parts = firstLine.split('=');
    return {
      totalRows: parseInt(parts[0]) || 53,
      totalCols: parseInt(parts[1]) || 42,
    };
  }, [fmtText]);

  const fields = useMemo(() => {
    const lines = fmtText.split('\n').filter((l, i) => i > 0 && l.trim() && l.includes('='));
    return lines.map(line => {
      const parts = line.split('=');
      if (parts.length < 8) return null;
      return {
        startRow: parseInt(parts[0]),
        endRow: parseInt(parts[1]),
        startCol: parseInt(parts[2]),
        endCol: parseInt(parts[3]),
        type: parts[4],
        direction: parts[5],
        chars: parts[6],
        label: parts[8] || parts[7] || 'Field',
      } as OpticField;
    }).filter(Boolean) as OpticField[];
  }, [fmtText]);

  const rawData = useMemo(() => {
    return studentText.split('\n').filter(l => l.trim()).map(line => line.split('\t').map(p => p.trim()));
  }, [studentText]);

  const columnHeaders = useMemo(() => {
    if (rawData.length === 0) return [];
    if (hasHeader) return rawData[0];
    return rawData[0].map((_, i) => `Column ${i + 1}`);
  }, [rawData, hasHeader]);

  const students = useMemo(() => {
    const dataRows = hasHeader ? rawData.slice(1) : rawData;
    return dataRows.map(parts => {
      const data: StudentData = {};
      Object.entries(fieldMappings).forEach(([label, colIndex]) => {
        const idx = colIndex as number;
        data[label] = (parts as any[])[idx] !== undefined ? (parts as any[])[idx] : '';
      });
      return data;
    });
  }, [rawData, fieldMappings, hasHeader]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        
        if (data.length === 0) {
          showToast('Excel dosyası boş!', 'error');
          return;
        }

        // Convert to tab-separated text for the textarea
        const tsv = data.map(row => row.map(cell => cell === null || cell === undefined ? '' : cell).join('\t')).join('\n');
        setStudentText(tsv);
        setHasHeader(true);
        showToast('Excel verileri başarıyla yüklendi');
      } catch (err) {
        console.error('Excel okuma hatası:', err);
        showToast('Excel dosyası okunamadı!', 'error');
      }
    };
    reader.readAsBinaryString(file);
    // Reset input value to allow re-uploading same file
    e.target.value = '';
  };

  const handleFmtUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const text = evt.target?.result as string;
      if (text) {
        setFmtText(text);
      }
    };
    reader.readAsText(file);
  };

  const drawStudentOnDoc = (doc: jsPDF, student: StudentData) => {
    // Only print fields that have a valid mapping
    const activeFields = fields.filter(f => fieldMappings[f.label] !== undefined && fieldMappings[f.label] !== -1);

    activeFields.forEach(field => {
      let value = student[field.label] || '';
      const chars = field.chars;
      
      const numItems = field.direction === 'D' ? (field.endCol - field.startCol + 1) : (field.endRow - field.startRow + 1);

      // Right-align logic for Student Number
      if (field.label.toUpperCase().includes('ÖĞRENCİ NO') && value.length < numItems) {
        value = value.padStart(numItems, ' ');
      }

      for (let i = 0; i < numItems; i++) {
        const charToFill = value[i] || '';
        if (charToFill.trim() === '') continue;
        const charIndex = chars.indexOf(charToFill.toUpperCase());

        if (charIndex !== -1) {
          let bubbleX = 0;
          let bubbleY = 0;

          if (field.direction === 'D') {
            bubbleX = (field.startCol + i) * calibration.bubbleSpacing + calibration.offsetX;
            bubbleY = (field.startRow + charIndex) * calibration.bubbleSpacing + calibration.offsetY;
            
            // Print character 1/6" above the top row of the field, moved 1mm down closer to bubbles
            const textY = (field.startRow - 1) * calibration.bubbleSpacing + calibration.offsetY + 1;
            doc.setFontSize(7);
            doc.setTextColor(0, 0, 0);
            doc.text(charToFill.toUpperCase(), bubbleX, textY, { align: 'center' });
          } else {
            bubbleX = (field.startCol + charIndex) * calibration.bubbleSpacing + calibration.offsetX;
            bubbleY = (field.startRow + i) * calibration.bubbleSpacing + calibration.offsetY;
            
            // For horizontal fields, print text to the left
            const textX = (field.startCol - 1) * calibration.bubbleSpacing + calibration.offsetX;
            doc.setFontSize(7);
            doc.setTextColor(0, 0, 0);
            doc.text(charToFill.toUpperCase(), textX, bubbleY, { align: 'right', baseline: 'middle' });
          }

          doc.setFillColor(0, 0, 0);
          doc.ellipse(
            bubbleX, 
            bubbleY, 
            calibration.bubbleSize / 2, 
            calibration.bubbleSize / 2, 
            'F'
          );
        }
      }
    });
  };

  const loadTurkishFont = async (doc: jsPDF) => {
    try {
      const fontUrl = 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf';
      const response = await fetch(fontUrl);
      const buffer = await response.arrayBuffer();
      const base64 = btoa(new Uint8Array(buffer).reduce((data, byte) => data + String.fromCharCode(byte), ''));
      doc.addFileToVFS('Roboto-Medium.ttf', base64);
      doc.addFont('Roboto-Medium.ttf', 'Roboto', 'normal');
      doc.setFont('Roboto');
    } catch (e) {
      console.warn('Font loading failed, falling back to helvetica', e);
      doc.setFont('helvetica', 'bold');
    }
  };

  const generatePDF = async () => {
    setIsGenerating(true);
    try {
      const doc = new jsPDF({
        orientation: pageSize.width > pageSize.height ? 'landscape' : 'portrait',
        unit: 'mm',
        format: [pageSize.width * 10, pageSize.height * 10]
      });

      await loadTurkishFont(doc);

      students.forEach((student, index) => {
        if (index > 0) doc.addPage();
        drawStudentOnDoc(doc, student);
      });

      doc.save('optik_formlar.pdf');
    } catch (error) {
      console.error('PDF Oluşturma hatası:', error);
    } finally {
      setIsGenerating(false);
    }
  };

  const generatePDFsByClass = async () => {
    if (students.length === 0) return;
    setIsGenerating(true);
    try {
      const zip = new JSZip();
      
      // Group students by class and section
      const groups: Record<string, StudentData[]> = {};
      students.forEach(student => {
        const className = student['SINIF']?.toString().trim() || 'Belirsiz_Sinif';
        const sectionName = student['ŞUBE']?.toString().trim() || '';
        const groupKey = sectionName ? `${className}_${sectionName}` : className;
        
        if (!groups[groupKey]) groups[groupKey] = [];
        groups[groupKey].push(student);
      });

      const groupNames = Object.keys(groups);
      
      for (const groupName of groupNames) {
        const groupStudents = groups[groupName];
        const doc = new jsPDF({
          orientation: pageSize.width > pageSize.height ? 'landscape' : 'portrait',
          unit: 'mm',
          format: [pageSize.width * 10, pageSize.height * 10]
        });

        await loadTurkishFont(doc);

        groupStudents.forEach((student, index) => {
          if (index > 0) doc.addPage();
          drawStudentOnDoc(doc, student);
        });

        // Add PDF to ZIP
        const pdfData = doc.output('arraybuffer');
        zip.file(`optik_formlar_${groupName}.pdf`, pdfData);
      }

      // Generate ZIP file
      const content = await zip.generateAsync({ type: 'blob' });
      
      // Trigger download
      const link = document.createElement('a');
      link.href = URL.createObjectURL(content);
      link.download = 'optik_formlar_sinif_bazli.zip';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(link.href);

      showToast(`${groupNames.length} adet sınıf dosyası ZIP olarak oluşturuldu`);
    } catch (error) {
      console.error('Sınıf bazlı PDF oluşturma hatası:', error);
      showToast('PDF oluşturulurken hata oluştu', 'error');
    } finally {
      setIsGenerating(false);
    }
  };

  const printSample = async () => {
    if (students.length === 0) return;
    setIsGenerating(true);
    try {
      const doc = new jsPDF({
        orientation: pageSize.width > pageSize.height ? 'landscape' : 'portrait',
        unit: 'mm',
        format: [pageSize.width * 10, pageSize.height * 10]
      });

      await loadTurkishFont(doc);
      drawStudentOnDoc(doc, students[0]);

      // Trigger print dialog directly
      doc.autoPrint();
      window.open(doc.output('bloburl'), '_blank');
      showToast('Yazdırma penceresi açıldı');
    } catch (error) {
      console.error('Örnek yazdırma hatası:', error);
      showToast('Yazdırma sırasında hata oluştu', 'error');
    } finally {
      setIsGenerating(false);
    }
  };

  if (!isAuthorized) {
    return (
      <div className="min-h-screen bg-gray-100 flex items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-8 rounded-2xl shadow-2xl w-full max-w-md border border-gray-200"
        >
          <div className="flex flex-col items-center gap-6">
            <div className="w-16 h-16 bg-indigo-100 rounded-full flex items-center justify-center text-indigo-600">
              <Layout size={32} />
            </div>
            <div className="text-center">
              <h1 className="text-2xl font-bold text-gray-900">OptikForm Pro</h1>
              <p className="text-gray-500 mt-1">Lütfen devam etmek için parolayı girin</p>
            </div>
            
            <form onSubmit={handlePasswordSubmit} className="w-full space-y-4">
              <div className="space-y-2">
                <input
                  type="password"
                  value={passwordInput}
                  onChange={(e) => setPasswordInput(e.target.value)}
                  placeholder="Parola"
                  className={cn(
                    "w-full px-4 py-3 rounded-xl border outline-none transition-all",
                    passwordError ? "border-red-500 bg-red-50" : "border-gray-200 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-200"
                  )}
                  autoFocus
                />
                {passwordError && (
                  <p className="text-red-500 text-xs font-medium px-1">Hatalı parola, lütfen tekrar deneyin.</p>
                )}
              </div>
              <button
                type="submit"
                className="w-full py-3 bg-indigo-600 text-white rounded-xl font-bold hover:bg-indigo-700 transition-colors shadow-lg shadow-indigo-100 active:scale-[0.98]"
              >
                Giriş Yap
              </button>
            </form>

            <div className="bg-amber-50 border border-amber-200 p-5 rounded-xl w-full shadow-sm">
              <p className="text-amber-900 text-center leading-relaxed">
                <span className="text-sm font-medium block mb-1 opacity-80">Parola Almak İçin:</span>
                <span className="text-xl font-bold block tracking-tight">0532 650 24 63</span>
                <span className="text-xs mt-1 block opacity-70 italic">numaralı telefonu arayın</span>
              </p>
            </div>
            
            <p className="text-[10px] text-gray-400 uppercase tracking-widest font-bold">
              Güvenli Erişim Kontrolü
            </p>
          </div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-indigo-600 rounded-lg flex items-center justify-center text-white shadow-indigo-200 shadow-lg">
            <Layout size={24} />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight">OptikForm Pro</h1>
            <p className="text-xs text-gray-500 font-medium uppercase tracking-wider">Hassas OMR Ön Kodlama</p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <button
            onClick={generatePDF}
            disabled={isGenerating || students.length === 0}
            className={cn(
              "flex items-center gap-2 px-5 py-2.5 rounded-full font-semibold transition-all shadow-md",
              isGenerating 
                ? "bg-gray-100 text-gray-400 cursor-not-allowed" 
                : "bg-indigo-600 text-white hover:bg-indigo-700 active:scale-95 shadow-indigo-100"
            )}
          >
            {isGenerating ? (
              <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
            ) : (
              <Printer size={18} />
            )}
            {isGenerating ? "Oluşturuluyor..." : "PDF Yazdır"}
          </button>
        </div>
      </header>

      {/* Toast Notification */}
      <AnimatePresence>
        {toast && (
          <motion.div
            initial={{ opacity: 0, y: 50, x: '-50%' }}
            animate={{ opacity: 1, y: 0, x: '-50%' }}
            exit={{ opacity: 0, y: 20, x: '-50%' }}
            className={cn(
              "fixed bottom-8 left-1/2 z-[100] px-6 py-3 rounded-full shadow-2xl flex items-center gap-3 font-bold text-sm",
              toast.type === 'success' ? "bg-indigo-600 text-white" : "bg-red-600 text-white"
            )}
          >
            {toast.type === 'success' ? <CheckCircle2 size={18} /> : <AlertCircle size={18} />}
            {toast.message}
          </motion.div>
        )}
      </AnimatePresence>

      <main className="max-w-7xl mx-auto p-6 grid grid-cols-1 lg:grid-cols-12 gap-8">
        {/* Left Column: Configuration */}
        <div className="lg:col-span-8 space-y-6">
          {/* Tabs */}
          <div className="flex bg-white p-1 rounded-xl border border-gray-200 shadow-sm">
            {[
              { id: 'data', label: 'Öğrenci Verileri', icon: Users },
              { id: 'format', label: 'FMT Tanımı', icon: FileText },
              { id: 'settings', label: 'Kalibrasyon', icon: Settings },
              { id: 'preview', label: 'Ön İzleme', icon: Eye },
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id as any)}
                className={cn(
                  "flex-1 flex items-center justify-center gap-2 py-2.5 rounded-lg text-sm font-semibold transition-all",
                  activeTab === tab.id 
                    ? "bg-indigo-50 text-indigo-700 shadow-sm" 
                    : "text-gray-500 hover:text-gray-700 hover:bg-gray-50"
                )}
              >
                <tab.icon size={16} />
                {tab.label}
              </button>
            ))}
          </div>

          {/* Content Area */}
          <div className="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden min-h-[500px] flex flex-col">
            <AnimatePresence mode="wait">
              {activeTab === 'data' && (
                <motion.div
                  key="data"
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -10 }}
                  className="p-6 flex-1 flex flex-col gap-6"
                >
                  <div className="flex flex-col gap-4">
                    <div className="flex items-center justify-between">
                      <h2 className="font-bold text-lg flex items-center gap-2">
                        <Users size={20} className="text-indigo-600" />
                        Öğrenci Listesi ve Eşleştirme
                      </h2>
                      <div className="flex items-center gap-3">
                        <label className="flex items-center gap-2 text-xs font-semibold text-gray-500 cursor-pointer hover:text-indigo-600 transition-colors">
                          <input 
                            type="checkbox" 
                            checked={hasHeader} 
                            onChange={(e) => setHasHeader(e.target.checked)}
                            className="w-4 h-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500"
                          />
                          İlk satır başlıktır
                        </label>
                        <button
                          onClick={downloadSampleExcel}
                          className="flex items-center gap-2 px-3 py-1.5 bg-amber-50 text-amber-700 rounded-lg text-xs font-bold hover:bg-amber-100 transition-colors"
                        >
                          <Download size={14} />
                          Örnek Excel İndir
                        </button>
                        <label className="flex items-center gap-2 px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs font-bold cursor-pointer hover:bg-indigo-100 transition-colors">
                          <Upload size={14} />
                          Excel İçe Aktar
                          <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} className="hidden" />
                        </label>
                        <button 
                          onClick={() => setStudentText('')}
                          className="p-1.5 text-gray-400 hover:text-red-500 transition-colors"
                          title="Verileri Temizle"
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    </div>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="flex flex-col gap-2">
                        <div className="flex items-center justify-between">
                          <p className="text-sm font-bold text-gray-700">Veri Yapıştır (Sekme ile ayrılmış)</p>
                          <span className="text-[10px] text-gray-400 font-mono">{students.length} satır</span>
                        </div>
                        <textarea
                          value={studentText}
                          onChange={(e) => setStudentText(e.target.value)}
                          className="h-48 w-full p-3 font-mono text-xs bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none resize-none"
                          placeholder="Excel'den kopyalayıp buraya yapıştırın..."
                        />
                      </div>
                      
                      <div className="flex flex-col gap-2">
                        <p className="text-sm font-bold text-gray-700">Alan Eşleştirme</p>
                        <div className="bg-gray-50 border border-gray-200 rounded-xl p-3 h-48 overflow-y-auto space-y-2">
                          {fields.map((field) => {
                            const isMapped = fieldMappings[field.label] !== undefined && fieldMappings[field.label] !== -1;
                            return (
                              <div key={field.label} className={cn(
                                "flex items-center justify-between gap-2 p-2 rounded border transition-all shadow-sm",
                                isMapped ? "bg-indigo-50 border-indigo-200" : "bg-white border-gray-100"
                              )}>
                                <div className="flex flex-col min-w-0">
                                  <div className="flex items-center gap-2">
                                    <span className={cn("text-[10px] font-bold truncate", isMapped ? "text-indigo-700" : "text-gray-700")}>
                                      {field.label}
                                    </span>
                                    {isMapped && <CheckCircle2 size={10} className="text-indigo-500" />}
                                  </div>
                                  <span className="text-[8px] text-gray-400 font-mono">Sütunlar: {field.startCol}-{field.endCol}</span>
                                </div>
                                <select
                                  value={fieldMappings[field.label] ?? -1}
                                  onChange={(e) => setFieldMappings({ ...fieldMappings, [field.label]: parseInt(e.target.value) })}
                                  className="text-[10px] bg-white border border-gray-200 rounded px-1 py-1 outline-none focus:ring-1 focus:ring-indigo-500 min-w-[100px]"
                                >
                                  <option value={-1}>Yazdırma</option>
                                  {columnHeaders.map((header, i) => (
                                    <option key={i} value={i}>{header}</option>
                                  ))}
                                </select>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="flex-1 overflow-hidden flex flex-col border border-gray-200 rounded-xl bg-white">
                    <div className="bg-gray-50 border-b border-gray-200 px-4 py-2 flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <TableIcon size={14} className="text-gray-400" />
                        <span className="text-xs font-bold text-gray-500 uppercase tracking-wider">Veri Önizleme</span>
                      </div>
                      <span className="text-[10px] text-gray-400 italic">İlk 10 öğrenci gösteriliyor</span>
                    </div>
                    <div className="overflow-auto flex-1">
                      <table className="w-full text-left border-collapse min-w-full">
                        <thead className="sticky top-0 z-10">
                          <tr className="bg-white border-b border-gray-100 shadow-sm">
                            <th className="px-4 py-2 text-[10px] font-bold text-gray-400 uppercase w-10">#</th>
                            {fields.filter(f => fieldMappings[f.label] !== undefined && fieldMappings[f.label] !== -1).map(f => (
                              <th key={f.label} className="px-4 py-2 text-[10px] font-bold text-indigo-400 uppercase whitespace-nowrap">{f.label}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {students.length === 0 ? (
                            <tr>
                              <td colSpan={fields.length + 1} className="px-4 py-12 text-center text-gray-400 text-sm italic">
                                Veri yok. Excel'den yapıştırın veya dosya yükleyin.
                              </td>
                            </tr>
                          ) : (
                            students.slice(0, 10).map((s, i) => (
                              <tr key={i} className="border-b border-gray-50 hover:bg-gray-50/50 transition-colors">
                                <td className="px-4 py-2 text-[10px] font-mono text-gray-300">{i + 1}</td>
                                {fields.filter(f => fieldMappings[f.label] !== undefined && fieldMappings[f.label] !== -1).map(f => {
                                  let val = s[f.label] || '';
                                  const numCols = f.endCol - f.startCol + 1;
                                  const isStudentNo = f.label.toUpperCase().includes('ÖĞRENCİ NO');
                                  
                                  return (
                                    <td key={f.label} className="px-4 py-2 text-xs font-medium text-gray-600 whitespace-nowrap">
                                      {val ? (
                                        <div className={cn("flex", isStudentNo ? "justify-end" : "justify-start")}>
                                          <span className={cn(
                                            "px-1.5 py-0.5 rounded font-mono",
                                            isStudentNo ? "bg-amber-50 text-amber-700 border border-amber-100" : "bg-indigo-50 text-indigo-700"
                                          )}>
                                            {isStudentNo ? val.padStart(numCols, '•').replace(/•/g, ' ') : val}
                                            {isStudentNo && <span className="text-[8px] opacity-30 ml-1"> (Sağ)</span>}
                                          </span>
                                        </div>
                                      ) : (
                                        <span className="text-gray-300">-</span>
                                      )}
                                    </td>
                                  );
                                })}
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </motion.div>
              )}

              {activeTab === 'format' && (
                <motion.div
                  key="format"
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -10 }}
                  className="p-6 flex-1 flex flex-col gap-6"
                >
                  <div className="flex items-center justify-between">
                    <h2 className="font-bold text-lg flex items-center gap-2">
                      <FileText size={20} className="text-indigo-600" />
                      FMT Tanımı
                    </h2>
                    <div className="flex items-center gap-4">
                      <label className="flex items-center gap-2 px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs font-bold cursor-pointer hover:bg-indigo-100 transition-colors">
                        <Upload size={14} />
                        FMT İçe Aktar (.txt)
                        <input type="file" accept=".txt" onChange={handleFmtUpload} className="hidden" />
                      </label>
                      <button 
                        onClick={() => setFmtText(DEFAULT_FMT)}
                        className="text-xs font-bold text-indigo-600 hover:underline"
                      >
                        Varsayılana Sıfırla
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 flex-1 overflow-hidden">
                    {/* Visual Editor */}
                    <div className="flex flex-col gap-3 overflow-hidden">
                      <p className="text-sm font-bold text-gray-700">Görsel Alan Düzenleyici</p>
                      <div className="flex-1 overflow-y-auto border border-gray-200 rounded-xl bg-gray-50 p-2 space-y-2">
                        {fields.map((field, idx) => (
                          <div key={idx} className="bg-white p-3 rounded-lg border border-gray-100 shadow-sm space-y-3">
                            <div className="flex items-center justify-between">
                              <input 
                                value={field.label}
                                onChange={(e) => {
                                  const newFields = [...fields];
                                  newFields[idx].label = e.target.value;
                                  // Reconstruct FMT text
                                  setFmtText(newFields.map(f => `${f.startRow}=${f.endRow}=${f.startCol}=${f.endCol}=${f.type}=${f.direction}=${f.chars}=X2=${f.label}=`).join('\n'));
                                }}
                                className="font-bold text-sm text-indigo-600 bg-transparent border-b border-transparent hover:border-indigo-200 focus:border-indigo-500 outline-none px-1"
                              />
                              <span className="text-[10px] font-mono text-gray-400">Alan #{idx + 1}</span>
                            </div>
                            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                              <div className="space-y-1">
                                <label className="text-[9px] uppercase font-bold text-gray-400">Başlangıç Sütun</label>
                                <input 
                                  type="number"
                                  value={field.startCol}
                                  onChange={(e) => {
                                    const newFields = [...fields];
                                    newFields[idx].startCol = parseInt(e.target.value) || 0;
                                    setFmtText(newFields.map(f => `${f.startRow}=${f.endRow}=${f.startCol}=${f.endCol}=${f.type}=${f.direction}=${f.chars}=X2=${f.label}=`).join('\n'));
                                  }}
                                  className="w-full p-1 bg-gray-50 border border-gray-100 rounded text-xs font-mono"
                                />
                              </div>
                              <div className="space-y-1">
                                <label className="text-[9px] uppercase font-bold text-gray-400">Bitiş Sütun</label>
                                <input 
                                  type="number"
                                  value={field.endCol}
                                  onChange={(e) => {
                                    const newFields = [...fields];
                                    newFields[idx].endCol = parseInt(e.target.value) || 0;
                                    setFmtText(newFields.map(f => `${f.startRow}=${f.endRow}=${f.startCol}=${f.endCol}=${f.type}=${f.direction}=${f.chars}=X2=${f.label}=`).join('\n'));
                                  }}
                                  className="w-full p-1 bg-gray-50 border border-gray-100 rounded text-xs font-mono"
                                />
                              </div>
                              <div className="space-y-1">
                                <label className="text-[9px] uppercase font-bold text-gray-400">Başlangıç Satır</label>
                                <input 
                                  type="number"
                                  value={field.startRow}
                                  onChange={(e) => {
                                    const newFields = [...fields];
                                    newFields[idx].startRow = parseInt(e.target.value) || 0;
                                    setFmtText(newFields.map(f => `${f.startRow}=${f.endRow}=${f.startCol}=${f.endCol}=${f.type}=${f.direction}=${f.chars}=X2=${f.label}=`).join('\n'));
                                  }}
                                  className="w-full p-1 bg-gray-50 border border-gray-100 rounded text-xs font-mono"
                                />
                              </div>
                              <div className="space-y-1">
                                <label className="text-[9px] uppercase font-bold text-gray-400">Bitiş Satır</label>
                                <input 
                                  type="number"
                                  value={field.endRow}
                                  onChange={(e) => {
                                    const newFields = [...fields];
                                    newFields[idx].endRow = parseInt(e.target.value) || 0;
                                    setFmtText(newFields.map(f => `${f.startRow}=${f.endRow}=${f.startCol}=${f.endCol}=${f.type}=${f.direction}=${f.chars}=X2=${f.label}=`).join('\n'));
                                  }}
                                  className="w-full p-1 bg-gray-50 border border-gray-100 rounded text-xs font-mono"
                                />
                              </div>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Raw Text Editor */}
                    <div className="flex flex-col gap-3 overflow-hidden">
                      <p className="text-sm font-bold text-gray-700">Ham FMT Metni</p>
                      <textarea
                        value={fmtText}
                        onChange={(e) => setFmtText(e.target.value)}
                        className="flex-1 w-full p-4 font-mono text-[10px] bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none resize-none leading-relaxed"
                      />
                    </div>
                  </div>
                </motion.div>
              )}

              {activeTab === 'settings' && (
                <motion.div
                  key="settings"
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -10 }}
                  className="p-6 space-y-8"
                >
                  <div className="space-y-6">
                    <div className="flex items-center justify-between">
                      <h2 className="font-bold text-lg flex items-center gap-2">
                        <Settings size={20} className="text-indigo-600" />
                        Kalibrasyon Ayarları
                      </h2>
                      <button
                        onClick={generateTestPage}
                        className="flex items-center gap-2 px-4 py-2 bg-amber-50 text-amber-700 rounded-lg text-sm font-bold hover:bg-amber-100 transition-colors"
                      >
                        <Printer size={16} />
                        Milimetrik Test Sayfası Yazdır
                      </button>
                    </div>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-4">
                        <label className="block">
                          <span className="text-sm font-bold text-gray-700 block mb-2">X Ofset (mm)</span>
                          <input
                            type="number"
                            step="0.001"
                            value={calibration.offsetX}
                            onChange={(e) => setCalibration({ ...calibration, offsetX: parseFloat(e.target.value) })}
                            className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
                          />
                        </label>
                        <label className="block">
                          <span className="text-sm font-bold text-gray-700 block mb-2">Y Ofset (mm)</span>
                          <input
                            type="number"
                            step="0.001"
                            value={calibration.offsetY}
                            onChange={(e) => setCalibration({ ...calibration, offsetY: parseFloat(e.target.value) })}
                            className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
                          />
                        </label>
                      </div>
                      <div className="space-y-4">
                        <label className="block">
                          <span className="text-sm font-bold text-gray-700 block mb-2">Baloncuk Aralığı (mm)</span>
                          <div className="relative">
                            <input
                              type="number"
                              step="0.000001"
                              value={calibration.bubbleSpacing}
                              onChange={(e) => setCalibration({ ...calibration, bubbleSpacing: parseFloat(e.target.value) })}
                              className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
                            />
                            <span className="absolute right-3 top-2 text-[10px] text-gray-400 font-mono">
                              1/6" ≈ 4.233mm
                            </span>
                          </div>
                        </label>
                        <label className="block">
                          <span className="text-sm font-bold text-gray-700 block mb-2">Baloncuk Boyutu (mm)</span>
                          <input
                            type="number"
                            step="0.1"
                            value={calibration.bubbleSize}
                            onChange={(e) => setCalibration({ ...calibration, bubbleSize: parseFloat(e.target.value) })}
                            className="w-full px-4 py-2 bg-gray-50 border border-gray-200 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
                          />
                        </label>
                      </div>
                    </div>
                  </div>

                  <div className="bg-indigo-50 p-4 rounded-xl border border-indigo-100 flex gap-3">
                    <Info size={20} className="text-indigo-600 shrink-0" />
                    <p className="text-xs text-indigo-800 leading-relaxed">
                      <strong>Hassasiyet İpucu:</strong> Ofsetleri küçük adımlarla (0.1mm veya daha az) ayarlayarak basılı baloncukları fiziksel formlarınızla mükemmel şekilde hizalayın. Baloncuk aralığı varsayılan olarak endüstri standardı olan 1/6 inç'e ayarlanmıştır.
                    </p>
                  </div>

                  {/* Profile Management Section */}
                  <div className="pt-6 border-t border-gray-100">
                    <div className="flex items-center justify-between mb-4">
                      <h3 className="font-bold text-gray-900 flex items-center gap-2">
                        <Bookmark size={18} className="text-indigo-600" />
                        Kayıtlı Optik Form Profilleri
                      </h3>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                      {/* Save Current Settings */}
                      <div className="bg-gray-50 p-4 rounded-xl border border-gray-200">
                        <p className="text-sm font-bold text-gray-700 mb-3">Mevcut Ayarları Kaydet</p>
                        <div className="flex gap-2">
                          <input
                            type="text"
                            placeholder="Profil adı (örn: 8. Sınıf LGS)"
                            value={profileName}
                            onChange={(e) => setProfileName(e.target.value)}
                            className="flex-1 px-3 py-2 bg-white border border-gray-200 rounded-lg text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
                          />
                          <button
                            onClick={saveProfile}
                            disabled={!profileName.trim()}
                            className="px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-bold hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-2"
                          >
                            <Save size={16} />
                            Kaydet
                          </button>
                        </div>
                      </div>

                      {/* List Profiles */}
                      <div className="space-y-2 max-h-[200px] overflow-y-auto pr-2">
                        {profiles.length === 0 ? (
                          <div className="text-center py-8 bg-gray-50 rounded-xl border border-dashed border-gray-200">
                            <p className="text-xs text-gray-400">Henüz kayıtlı profil yok.</p>
                          </div>
                        ) : (
                          profiles.map((profile) => (
                            <div 
                              key={profile.id}
                              className="group flex items-center justify-between p-3 bg-white border border-gray-200 rounded-xl hover:border-indigo-300 hover:shadow-sm transition-all"
                            >
                              <div className="flex items-center gap-3 min-w-0">
                                <div className="w-8 h-8 bg-indigo-50 rounded-lg flex items-center justify-center text-indigo-600">
                                  <FileText size={16} />
                                </div>
                                <div className="flex flex-col min-w-0">
                                  <span className="text-sm font-bold text-gray-700 truncate">{profile.name}</span>
                                  <span className="text-[10px] text-gray-400">
                                    {new Date(profile.createdAt).toLocaleDateString('tr-TR')}
                                  </span>
                                </div>
                              </div>
                              <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                                <button
                                  onClick={() => loadProfile(profile)}
                                  className="p-1.5 text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors"
                                  title="Yükle"
                                >
                                  <History size={16} />
                                </button>
                                <button
                                  onClick={() => deleteProfile(profile.id)}
                                  className="p-1.5 text-red-500 hover:bg-red-50 rounded-lg transition-colors"
                                  title="Sil"
                                >
                                  <Trash2 size={16} />
                                </button>
                              </div>
                            </div>
                          ))
                        )}
                      </div>
                    </div>
                  </div>
                </motion.div>
              )}
              {activeTab === 'preview' && (
                <motion.div
                  key="preview"
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 0.95 }}
                  className="p-6 flex-1 flex flex-col gap-4 overflow-hidden"
                >
                  <div className="flex items-center justify-between">
                    <div className="flex flex-col">
                      <h2 className="font-bold text-lg flex items-center gap-2">
                        <Eye size={20} className="text-indigo-600" />
                        Form Ön İzleme
                      </h2>
                      <p className="text-xs text-gray-500 italic">
                        İlk öğrencinin verileriyle örnek görünüm
                      </p>
                    </div>
                    
                    <button
                      onClick={printSample}
                      disabled={isGenerating || students.length === 0}
                      className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg font-bold text-sm hover:bg-indigo-700 transition-all shadow-md active:scale-95"
                    >
                      <Printer size={16} />
                      Örnek Yazdır
                    </button>
                  </div>
                  
                  <div className="flex-1 bg-gray-100 rounded-xl border border-gray-200 p-4 overflow-auto flex justify-center">
                    <div 
                      className="bg-white shadow-2xl relative origin-top"
                      style={{
                        width: `${pageSize.width * 10}mm`,
                        height: `${pageSize.height * 10}mm`,
                        minWidth: `${pageSize.width * 10}mm`,
                        minHeight: `${pageSize.height * 10}mm`,
                        transform: 'scale(0.8)'
                      }}
                    >
                      {/* Grid/Ruler (Optional) */}
                      <div className="absolute inset-0 pointer-events-none opacity-[0.03]" style={{ backgroundImage: 'radial-gradient(#000 1px, transparent 1px)', backgroundSize: '10mm 10mm' }} />
                      
                      {/* Render Fields */}
                      {fields.filter(f => fieldMappings[f.label] !== undefined && fieldMappings[f.label] !== -1).map((field, fIdx) => {
                        const student = students[0] || {};
                        let value = student[field.label] || '';
                        const numItems = field.direction === 'D' ? (field.endCol - field.startCol + 1) : (field.endRow - field.startRow + 1);
                        if (field.label.toUpperCase().includes('ÖĞRENCİ NO') && value.length < numItems) {
                          value = value.padStart(numItems, ' ');
                        }

                        return (
                          <div key={fIdx}>
                            {Array.from({ length: numItems }).map((_, i) => {
                              const charToFill = value[i] || '';
                              const charIndex = charToFill.trim() === '' ? -1 : field.chars.indexOf(charToFill.toUpperCase());
                              
                              // Header Text
                              let textX = 0;
                              let textY = 0;
                              if (field.direction === 'D') {
                                textX = (field.startCol + i) * calibration.bubbleSpacing + calibration.offsetX;
                                textY = (field.startRow - 1) * calibration.bubbleSpacing + calibration.offsetY + 1;
                              } else {
                                textX = (field.startCol - 1) * calibration.bubbleSpacing + calibration.offsetX;
                                textY = (field.startRow + i) * calibration.bubbleSpacing + calibration.offsetY;
                              }

                              return (
                                <React.Fragment key={i}>
                                  {charToFill.trim() !== '' && (
                                    <div 
                                      className="absolute text-[6px] font-bold text-black flex items-center justify-center"
                                      style={{
                                        left: `${textX}mm`,
                                        top: `${textY}mm`,
                                        width: `${calibration.bubbleSize}mm`,
                                        height: `${calibration.bubbleSize}mm`,
                                        transform: 'translate(-50%, -50%)'
                                      }}
                                    >
                                      {charToFill.toUpperCase()}
                                    </div>
                                  )}
                                  
                                  {/* Bubbles for this column */}
                                  {field.chars.split('').map((char, charIdx) => {
                                    let bX = 0;
                                    let bY = 0;
                                    if (field.direction === 'D') {
                                      bX = (field.startCol + i) * calibration.bubbleSpacing + calibration.offsetX;
                                      bY = (field.startRow + charIdx) * calibration.bubbleSpacing + calibration.offsetY;
                                    } else {
                                      bX = (field.startCol + charIdx) * calibration.bubbleSpacing + calibration.offsetX;
                                      bY = (field.startRow + i) * calibration.bubbleSpacing + calibration.offsetY;
                                    }

                                    const isFilled = charIdx === charIndex;

                                    return (
                                      <div 
                                        key={charIdx}
                                        className={cn(
                                          "absolute rounded-full border border-gray-300",
                                          isFilled ? "bg-black border-black" : "bg-transparent"
                                        )}
                                        style={{
                                          left: `${bX}mm`,
                                          top: `${bY}mm`,
                                          width: `${calibration.bubbleSize}mm`,
                                          height: `${calibration.bubbleSize}mm`,
                                          transform: 'translate(-50%, -50%)'
                                        }}
                                      />
                                    );
                                  })}
                                </React.Fragment>
                              );
                            })}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </div>

        {/* Right Column: Preview & Summary */}
        <div className="lg:col-span-4 space-y-6">
          <div className="bg-white rounded-2xl border border-gray-200 shadow-sm p-6">
            <h3 className="font-bold text-gray-900 mb-4 flex items-center gap-2">
              <Layout size={18} className="text-indigo-600" />
              Form Özeti
            </h3>
            
            <div className="space-y-4">
              <div className="grid grid-cols-2 gap-3">
                <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[10px] text-gray-400 uppercase font-bold tracking-widest mb-1">Genişlik (cm)</p>
                  <input 
                    type="number" 
                    step="0.1"
                    value={pageSize.width}
                    onChange={(e) => setPageSize({ ...pageSize, width: parseFloat(e.target.value) || 0 })}
                    className="w-full bg-transparent font-bold text-indigo-600 outline-none"
                  />
                </div>
                <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[10px] text-gray-400 uppercase font-bold tracking-widest mb-1">Yükseklik (cm)</p>
                  <input 
                    type="number" 
                    step="0.1"
                    value={pageSize.height}
                    onChange={(e) => setPageSize({ ...pageSize, height: parseFloat(e.target.value) || 0 })}
                    className="w-full bg-transparent font-bold text-indigo-600 outline-none"
                  />
                </div>
              </div>

              <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                <p className="text-[10px] text-gray-400 uppercase font-bold tracking-widest mb-1">Tespit Edilen Alanlar</p>
                <div className="flex flex-wrap gap-1.5">
                  {fields.map((f, i) => (
                    <span key={i} className="px-2 py-0.5 bg-white border border-gray-200 rounded text-[10px] font-semibold text-gray-600">
                      {f.label}
                    </span>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[10px] text-gray-400 uppercase font-bold tracking-widest mb-1">Öğrenciler</p>
                  <p className="text-xl font-bold text-indigo-600">{students.length}</p>
                </div>
                <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                  <p className="text-[10px] text-gray-400 uppercase font-bold tracking-widest mb-1">Sayfalar</p>
                  <p className="text-xl font-bold text-indigo-600">{students.length}</p>
                </div>
              </div>
            </div>
          </div>

          <div className="bg-[#151619] rounded-2xl shadow-xl p-6 text-white overflow-hidden relative group">
            <div className="absolute top-0 right-0 p-4 opacity-10 group-hover:opacity-20 transition-opacity">
              <Printer size={120} />
            </div>
            
            <h3 className="font-bold text-lg mb-2 relative z-10">Yazdırmaya Hazır mısınız?</h3>
            <p className="text-sm text-gray-400 mb-6 relative z-10">
              Kalibrasyonun doğru olması için yazıcınızın "Gerçek Boyut" (%100 ölçek) olarak ayarlandığından emin olun.
            </p>

            <div className="space-y-3 relative z-10">
              <button
                onClick={generatePDF}
                disabled={isGenerating || students.length === 0}
                className="w-full py-3 bg-white text-black rounded-xl font-bold flex items-center justify-center gap-2 hover:bg-gray-100 transition-colors"
              >
                <Download size={18} />
                Tümünü PDF İndir
              </button>
              
              <button
                onClick={generatePDFsByClass}
                disabled={isGenerating || students.length === 0}
                className="w-full py-3 bg-indigo-600 text-white rounded-xl font-bold flex items-center justify-center gap-2 hover:bg-indigo-700 transition-colors border border-indigo-500"
              >
                <Users size={18} />
                Sınıf Sınıf Ayrı PDF İndir
              </button>
            </div>
          </div>

          <div className="bg-white rounded-2xl border border-gray-200 shadow-sm p-6">
            <h3 className="font-bold text-gray-900 mb-4 flex items-center gap-2">
              <AlertCircle size={18} className="text-amber-500" />
              Talimatlar
            </h3>
            <ul className="space-y-3 text-xs text-gray-600">
              <li className="flex gap-2">
                <div className="w-4 h-4 rounded-full bg-amber-50 text-amber-600 flex items-center justify-center shrink-0 font-bold">1</div>
                <span>Öğrenci verilerini <strong>Öğrenci Verileri</strong> sekmesine yapıştırın.</span>
              </li>
              <li className="flex gap-2">
                <div className="w-4 h-4 rounded-full bg-amber-50 text-amber-600 flex items-center justify-center shrink-0 font-bold">2</div>
                <span><strong>FMT Tanımı</strong>'nın form düzeninizle eşleştiğini doğrulayın.</span>
              </li>
              <li className="flex gap-2">
                <div className="w-4 h-4 rounded-full bg-amber-50 text-amber-600 flex items-center justify-center shrink-0 font-bold">3</div>
                <span>Baloncukları fiziksel kağıdınızla hizalamak için <strong>Kalibrasyon</strong>'u kullanın.</span>
              </li>
              <li className="flex gap-2">
                <div className="w-4 h-4 rounded-full bg-amber-50 text-amber-600 flex items-center justify-center shrink-0 font-bold">4</div>
                <span>Bir test sayfası yazdırın ve gerekirse ofsetleri ölçün.</span>
              </li>
            </ul>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="max-w-7xl mx-auto px-6 py-8 border-t border-gray-200 mt-12 flex flex-col md:flex-row justify-between items-center gap-4">
        <div className="flex items-center gap-2 text-gray-400 text-sm">
          <CheckCircle2 size={16} className="text-green-500" />
          <span>Yüksek Hassasiyetli OMR Motoru Aktif</span>
        </div>
        <div className="text-gray-400 text-xs font-mono">
          Tolerans: 0.001mm | Aralık: 1/6"
        </div>
      </footer>
    </div>
  );
}
