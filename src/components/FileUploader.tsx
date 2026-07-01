import React, { useState, useRef } from "react";
import { motion, AnimatePresence } from "motion/react";
import { UploadCloud, FileSpreadsheet, FileText, Sparkles, AlertCircle } from "lucide-react";
import { ParsedDocument } from "../types";
import * as XLSX from "xlsx";

// Helper function for client-side parsing of documents
const parseFileClientSide = async (file: File): Promise<ParsedDocument> => {
  const extension = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
  
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const arrayBuffer = e.target?.result as ArrayBuffer;
        let extractedText = "";
        let visualContent: any = null;

        switch (extension) {
          case ".xlsx":
          case ".xls": {
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });
            const sheets: any[] = [];
            
            workbook.SheetNames.forEach(sheetName => {
              const worksheet = workbook.Sheets[sheetName];
              const grid = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
              if (grid.length > 0) {
                sheets.push({
                  sheetName,
                  grid: grid.slice(0, 1000), // Cap grid rows to protect UI rendering limits
                  rowCount: grid.length,
                  columnCount: grid[0]?.length || 0,
                });
              }
            });
            
            visualContent = { sheets };

            // Convert cells to CSV layout for context
            const csvLines: string[] = [];
            sheets.forEach(sh => {
              csvLines.push(`--- HOJA: ${sh.sheetName} ---`);
              sh.grid.slice(0, 100).forEach((row: any) => {
                csvLines.push(row.map((cell: any) => cell === undefined || cell === null ? "" : String(cell)).join("\t"));
              });
            });
            extractedText = csvLines.join("\n");
            break;
          }

          case ".docx": {
            const mammothInstance = (window as any).mammoth;
            if (!mammothInstance) {
              throw new Error("La librería de procesamiento de Word (Mammoth.js) no está disponible en este momento. Verifique su conexión de red.");
            }

            const htmlResult = await mammothInstance.convertToHtml({ arrayBuffer });
            const rawTextResult = await mammothInstance.extractRawText({ arrayBuffer });
            
            extractedText = rawTextResult.value;
            visualContent = {
              html: htmlResult.value,
              messagesCount: htmlResult.messages.length,
            };
            break;
          }

          case ".doc": {
            const decoder = new TextDecoder("utf-8");
            const str = decoder.decode(arrayBuffer);
            
            // Regex to extract readable blocks: letters, digits, spaces, common punctuation, Spanish accents/questions
            const matchedBlocks = str.match(/[\x20-\x7E\xA0-\xFF\s\u00C0-\u00FF\u0152\u0153\u0178]{15,}/g);
            
            if (!matchedBlocks || matchedBlocks.length === 0) {
              extractedText = "No se pudo extraer texto estructurado directo de este archivo .doc binario heredado de Word. Por favor, intente con un formato .docx.";
            } else {
              extractedText = matchedBlocks
                .map(block => block.trim())
                .map(block => block.replace(/\s+/g, " "))
                .filter(block => block.length > 20 && !(/[\x7F-\x9F]/.test(block)))
                .slice(0, 150) // Limit length to avoid extreme trash blocks
                .join("\n\n");
            }
            
            visualContent = {
              html: `<div class="font-sans whitespace-pre-wrap leading-relaxed text-gray-800 p-4 bg-amber-50/40 rounded-lg border border-amber-100/50">
                <p class="text-xs font-semibold text-amber-800 mb-2 font-mono">⚠️ FORMATO BINARIO DE WORD (.doc) DETECTADO</p>
                ${extractedText.replace(/\n\n/g, "</p><p class='mt-4'>")}
              </div>`,
              excerpt: extractedText,
            };
            break;
          }

          case ".pdf": {
            const pdfjsLib = (window as any).pdfjsLib;
            if (!pdfjsLib) {
              throw new Error("La librería de procesamiento de PDF (PDF.js) no está disponible en este momento. Verifique su conexión de red.");
            }

            // Set up worker
            pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js";

            const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) });
            const pdf = await loadingTask.promise;
            
            let pdfText = "";
            for (let i = 1; i <= pdf.numPages; i++) {
              const page = await pdf.getPage(i);
              const textContent = await page.getTextContent();
              const pageText = textContent.items
                .map((item: any) => item.str)
                .join(" ");
              pdfText += pageText + "\n";
            }
            
            extractedText = pdfText;
            visualContent = {
              pages: pdf.numPages,
              version: "1.0",
              info: {},
              excerpt: extractedText.trim().substring(0, 10000),
            };
            break;
          }

          default:
            throw new Error(`Extensión de archivo no soportada: ${extension}`);
        }

        // Clip text if it exceeds token safety limits
        let clipped = false;
        const wordLimit = 80000;
        const rawWords = extractedText.split(/\s+/);
        if (rawWords.length > wordLimit) {
          extractedText = rawWords.slice(0, wordLimit).join(" ");
          clipped = true;
        }

        resolve({
          success: true,
          filename: file.name,
          size: file.size,
          extension,
          visualContent,
          extractedText,
          clipped
        });

      } catch (err) {
        reject(err);
      }
    };
    
    reader.onerror = (err) => reject(new Error("Error al leer el archivo en disco."));
    reader.readAsArrayBuffer(file);
  });
};

interface FileUploaderProps {
  onFileUploaded: (doc: ParsedDocument) => void;
  onError: (msg: string) => void;
}

export default function FileUploader({ onFileUploaded, onError }: FileUploaderProps) {
  const [isDragActive, setIsDragActive] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [loadingPhrase, setLoadingPhrase] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);

  const loadingPhrases = [
    "Leyendo estructura del archivo...",
    "Extrayendo texto de bloques binarios...",
    "Conversión de codificación a texto plano...",
    "Generando análisis intelectual...",
    "Estructurando tablas y metadatos...",
    "Compilando resumen final...",
  ];

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setIsDragActive(true);
    } else if (e.type === "dragleave") {
      setIsDragActive(false);
    }
  };

  const uploadFile = async (file: File) => {
    const validExtensions = [".pdf", ".doc", ".docx", ".xls", ".xlsx"];
    const extension = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
    
    if (!validExtensions.includes(extension)) {
      onError(`El archivo "${file.name}" no está soportado. Formatos válidos: PDF, DOC, DOCX, XLS, XLSX`);
      return;
    }

    if (file.size > 25 * 1024 * 1024) {
      onError("El archivo es demasiado grande. El límite es de 25 MB.");
      return;
    }

    setIsUploading(true);
    setUploadProgress(10);
    setLoadingPhrase(loadingPhrases[0]);

    // Fast progressive bar simulator while parsing in client
    const interval = setInterval(() => {
      setUploadProgress(prev => {
        if (prev >= 90) {
          clearInterval(interval);
          return 90;
        }
        const step = Math.floor(Math.random() * 8) + 4;
        const phraseIdx = Math.floor((prev / 90) * loadingPhrases.length);
        if (loadingPhrases[phraseIdx]) {
          setLoadingPhrase(loadingPhrases[phraseIdx]);
        }
        return prev + step;
      });
    }, 150);

    try {
      // Parse file client-side directly
      const parsedData = await parseFileClientSide(file);

      clearInterval(interval);
      setUploadProgress(100);
      setLoadingPhrase("Análisis completado con éxito!");
      
      // Short delay for visual completion feel
      setTimeout(() => {
        onFileUploaded(parsedData);
        setIsUploading(false);
        setUploadProgress(0);
      }, 500);

    } catch (err: any) {
      clearInterval(interval);
      setIsUploading(false);
      setUploadProgress(0);
      onError(err.message || "Error al subir y procesar el documento.");
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      uploadFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      uploadFile(e.target.files[0]);
    }
  };

  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  return (
    <div className="w-full max-w-3xl mx-auto" id="file-uploader-container">
      <AnimatePresence mode="wait">
        {!isUploading ? (
          <motion.div
            key="dropzone"
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.95 }}
            transition={{ duration: 0.3 }}
            onDragEnter={handleDrag}
            onDragOver={handleDrag}
            onDragLeave={handleDrag}
            onDrop={handleDrop}
            onClick={triggerFileSelect}
            className={`relative flex flex-col items-center justify-center border-2 border-dashed rounded-3xl p-12 text-center cursor-pointer transition-all duration-300 min-h-[300px] select-none group
              ${isDragActive 
                ? "border-black bg-gray-50/80 scale-[1.01] shadow-md" 
                : "border-gray-200 bg-white hover:border-gray-400 hover:shadow-md"
              }`}
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileChange}
              accept=".pdf,.doc,.docx,.xls,.xlsx"
              className="hidden"
            />

            <motion.div
              animate={isDragActive ? { y: -8, scale: 1.05 } : { y: 0, scale: 1 }}
              transition={{ type: "spring", stiffness: 300, damping: 20 }}
              className={`w-16 h-16 rounded-full flex items-center justify-center mb-5 transition-all
                ${isDragActive 
                  ? "bg-black text-white" 
                  : "bg-gray-100 text-gray-400 group-hover:text-black group-hover:bg-gray-200"
                }`}
            >
              <UploadCloud className="w-7 h-7" />
            </motion.div>

            <h3 className="text-xl font-semibold text-gray-900 tracking-tight mb-2">
              {isDragActive ? "Soltá el archivo acá" : "Arrastrá tus documentos aquí"}
            </h3>
            
            <p className="text-xs text-gray-400 mt-1 max-w-sm leading-relaxed mb-6">
              Soporta formatos <span className="font-mono">.pdf</span>, <span className="font-mono">.doc / .docx</span> y planillas <span className="font-mono">.xls / .xlsx</span> de hasta 25MB.
            </p>

            <span className="inline-flex items-center gap-1.5 px-5 py-2.5 bg-black hover:bg-gray-850 text-white text-xs font-bold rounded-full shadow-sm transition-all duration-200 uppercase tracking-widest">
              Seleccionar archivo
            </span>

            {/* Formats info pills */}
            <div className="flex flex-wrap justify-center gap-3 mt-8 select-none">
              <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-red-50 text-red-600 text-[10px] font-bold border border-red-100 uppercase tracking-widest">
                <span className="w-1.5 h-1.5 rounded-full bg-red-500" /> PDF
              </span>
              <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-zinc-100 text-zinc-700 text-[10px] font-bold border border-zinc-200 uppercase tracking-widest">
                <span className="w-1.5 h-1.5 rounded-full bg-zinc-500" /> .doc / .docx
              </span>
              <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-green-50 text-green-600 text-[10px] font-bold border border-green-100 uppercase tracking-widest">
                <span className="w-1.5 h-1.5 rounded-full bg-green-500" /> .xls / .xlsx
              </span>
            </div>
          </motion.div>
        ) : (
          <motion.div
            key="loader"
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            exit={{ opacity: 0, y: -15 }}
            transition={{ duration: 0.25 }}
            className="flex flex-col items-center justify-center bg-white border border-gray-200 rounded-3xl p-12 text-center min-h-[300px] shadow-sm"
          >
            {/* Spinning pulse loader */}
            <div className="relative w-16 h-16 mb-5">
              <div className="absolute inset-0 rounded-full border-4 border-gray-100" />
              <motion.div
                animate={{ rotate: 360 }}
                transition={{ duration: 1, repeat: Infinity, ease: "linear" }}
                className="absolute inset-0 rounded-full border-4 border-black border-t-transparent"
              />
            </div>

            <h3 className="text-base font-semibold text-gray-900 mb-1">
              {loadingPhrase}
            </h3>
            
            <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest mb-6">
              Progreso: {uploadProgress}%
            </p>

            {/* Progress line */}
            <div className="w-full max-w-xs bg-gray-100 rounded-full h-1 overflow-hidden">
              <motion.div
                initial={{ width: "0%" }}
                animate={{ width: `${uploadProgress}%` }}
                transition={{ duration: 0.2 }}
                className="bg-black h-1 rounded-full"
              />
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
