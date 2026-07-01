import React, { useState, useEffect } from "react";
import { motion, AnimatePresence } from "motion/react";
import { 
  FileText, 
  AlertTriangle, 
  ShieldCheck, 
  X, 
  Download, 
  FileEdit, 
  Grid, 
  Layers, 
  CheckCircle, 
  FileSpreadsheet, 
  Sparkles, 
  HelpCircle,
  Clock,
  Briefcase,
  ChevronRight,
  Monitor
} from "lucide-react";
import FileUploader from "./components/FileUploader";
import DocPreview from "./components/DocPreview";
import { ParsedDocument } from "./types";

export default function App() {
  const [hasEntered, setHasEntered] = useState<boolean>(false);
  const [activeDoc, setActiveDoc] = useState<ParsedDocument | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [currentTime, setCurrentTime] = useState<string>("");

  // Update clock
  useEffect(() => {
    const updateTime = () => {
      const now = new Date();
      setCurrentTime(now.toLocaleTimeString("es-AR", { hour: "2-digit", minute: "2-digit", second: "2-digit" }));
    };
    updateTime();
    const interval = setInterval(updateTime, 1000);
    return () => clearInterval(interval);
  }, []);

  // Dismiss error after 6 seconds auto safety
  useEffect(() => {
    if (errorMsg) {
      const timer = setTimeout(() => {
        setErrorMsg(null);
      }, 6500);
      return () => clearTimeout(timer);
    }
  }, [errorMsg]);

  const handleFileUploaded = (doc: ParsedDocument) => {
    setActiveDoc(doc);
    setErrorMsg(null);
  };

  const handleError = (msg: string) => {
    setErrorMsg(msg);
  };

  const handleClearDocument = () => {
    setActiveDoc(null);
    setErrorMsg(null);
  };

  // Create an empty Elor_TT spreadsheet template instantly
  const handleCreateEmptySpreadsheet = () => {
    const emptyDoc: ParsedDocument = {
      success: true,
      filename: "Nueva_Planilla_Elor_TT.xlsx",
      size: 1024 * 5,
      extension: ".xlsx",
      visualContent: {
        sheets: [
          {
            sheetName: "Hoja Principal",
            grid: Array.from({ length: 45 }).map(() => Array.from({ length: 16 }).map(() => "")),
            rowCount: 45,
            columnCount: 16
          }
        ]
      },
      extractedText: "",
      clipped: false
    };
    setActiveDoc(emptyDoc);
  };

  // Create an empty Elor_TT Word/Doc template instantly
  const handleCreateEmptyDocument = () => {
    const emptyDoc: ParsedDocument = {
      success: true,
      filename: "Nuevo_Documento_Elor_TT.docx",
      size: 1024 * 3,
      extension: ".docx",
      visualContent: {
        html: ""
      },
      extractedText: "",
      clipped: false
    };
    setActiveDoc(emptyDoc);
  };

  // Create an empty Elor_TT PDF template instantly
  const handleCreateEmptyPdf = () => {
    const emptyDoc: ParsedDocument = {
      success: true,
      filename: "Nuevo_Documento_Elor_TT.pdf",
      size: 1024 * 4,
      extension: ".pdf",
      visualContent: {
        pages: 1,
        excerpt: ""
      },
      extractedText: "",
      clipped: false
    };
    setActiveDoc(emptyDoc);
  };

  return (
    <div className="min-h-screen bg-[#0F111A] text-[#E4E6EB] flex flex-col font-sans transition-colors duration-300 relative overflow-x-hidden selection:bg-[#343A4A] selection:text-white">
      
      {/* Heavy Industrial Graphic Rings (No blue, pure metallic slate and charcoal) */}
      <div className="absolute top-[-300px] right-[-300px] w-[700px] h-[700px] rounded-full border-[5px] border-zinc-900/40 pointer-events-none select-none z-0" />
      <div className="absolute top-[-250px] right-[-250px] w-[600px] h-[600px] rounded-full border-[2px] border-dashed border-zinc-800/20 pointer-events-none select-none z-0" />
      <div className="absolute bottom-[-200px] left-[-200px] w-[500px] h-[500px] rounded-full border-[4px] border-zinc-900/30 pointer-events-none select-none z-0" />
      <div className="absolute bottom-[-150px] left-[-150px] w-[400px] h-[400px] rounded-full border-[1px] border-dashed border-zinc-800/30 pointer-events-none select-none z-0" />

      {/* Global Header */}
      <header className="bg-[#161925] border-b border-zinc-800 sticky top-0 z-30 select-none shadow-xl">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-18 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-full bg-[#242836] border-2 border-zinc-700 flex items-center justify-center text-white font-black text-sm shadow-md">
              E_T
            </div>
            <div>
              <h1 className="text-base font-bold text-white tracking-wide flex items-center gap-2">
                Elor_TT <span className="font-semibold text-[9px] text-[#A3A8B8] bg-[#2E3344] px-2 py-0.5 rounded-full border border-zinc-700 uppercase tracking-widest font-mono scale-90">Nethels Industries</span>
              </h1>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="hidden md:flex items-center gap-2 bg-[#0F111A] border border-zinc-800 px-3 py-1.5 rounded-full text-zinc-350">
              <Clock className="w-3.5 h-3.5 text-zinc-400" />
              <span className="text-xs font-mono tracking-widest font-bold text-zinc-300">{currentTime || "--:--:--"}</span>
            </div>
            <div className="bg-[#1E2230] px-3 py-1.5 rounded-full border border-zinc-700 text-[10px] text-zinc-400 font-mono tracking-widest uppercase font-bold hidden sm:block">
              Nethels OS 2026
            </div>
          </div>
        </div>
      </header>

      {/* Main Workspace Frame container */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 flex flex-col z-10">
        
        {/* Toast / Alert banner notification */}
        <AnimatePresence>
          {errorMsg && (
            <motion.div
              initial={{ opacity: 0, y: -20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="bg-[#2C161B] border border-red-900/80 text-red-200 text-xs rounded-2xl p-4 mb-6 shadow-2xl flex items-start gap-3 relative backdrop-blur-md"
            >
              <AlertTriangle className="w-4 h-4 text-red-500 shrink-0 mt-0.5 animate-bounce" />
              <div className="flex-1">
                <h5 className="font-bold text-red-300 mb-0.5">Operación interrumpida</h5>
                <p className="leading-relaxed text-red-200/95">{errorMsg}</p>
              </div>
              <button
                onClick={() => setErrorMsg(null)}
                className="p-1 text-red-400 hover:text-red-100 transition rounded"
              >
                <X className="w-4 h-4" />
              </button>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Dynamic Display Router Stage */}
        <AnimatePresence mode="wait">
          {!hasEntered ? (
            /* INITIAL STARTUP / SPLASH SCREEN - PANTALLA DE INICIO */
            <motion.div
              key="splash"
              initial={{ opacity: 0, scale: 0.96 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              transition={{ duration: 0.4 }}
              className="flex-1 flex flex-col items-center justify-center my-8 select-none"
            >
              {/* Heavy, high-contrast, mechanical circular outer container */}
              <div className="w-full max-w-xl bg-[#161925] p-10 rounded-[3rem] border-4 border-zinc-700 shadow-2xl relative overflow-hidden flex flex-col items-center text-center">
                
                {/* Embedded Concentric Circles in Card */}
                <div className="absolute top-0 right-0 w-32 h-32 rounded-full border border-zinc-800 pointer-events-none" />
                <div className="absolute bottom-[-40px] left-[-40px] w-48 h-48 rounded-full border-2 border-dashed border-zinc-800 pointer-events-none" />

                {/* Circular indicator light console */}
                <div className="flex gap-2 mb-8 bg-[#0F111A] px-4 py-2 rounded-full border border-zinc-800">
                  <span className="w-2.5 h-2.5 rounded-full bg-zinc-600 animate-pulse" />
                  <span className="w-2.5 h-2.5 rounded-full bg-zinc-600" />
                  <span className="w-2.5 h-2.5 rounded-full bg-zinc-700" />
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-ping" />
                </div>

                {/* Large Center Circular Badge */}
                <div className="w-28 h-28 rounded-full bg-[#1F2436] border-4 border-zinc-600 flex items-center justify-center text-white font-black text-3xl shadow-lg mb-6 relative">
                  <div className="absolute inset-1 rounded-full border border-zinc-700/60 border-dashed animate-spin [animation-duration:12s]" />
                  E_T
                </div>

                {/* Identity Lockup */}
                <h2 className="text-4xl font-black text-white tracking-widest mb-1.5 uppercase font-mono">
                  Elor_TT
                </h2>
                <div className="text-xs text-zinc-400 uppercase tracking-[0.25em] font-bold font-mono mb-6">
                  Nethels Industries
                </div>

                <div className="w-full h-px bg-zinc-800 my-4" />

                {/* Double Contoured Enter Trigger Button */}
                <button
                  onClick={() => setHasEntered(true)}
                  className="mt-4 group relative inline-flex items-center justify-center p-0.5 mb-2 overflow-hidden text-sm font-bold text-white rounded-full bg-gradient-to-br from-zinc-600 to-zinc-800 hover:text-white focus:ring-4 focus:outline-none focus:ring-zinc-800 cursor-pointer"
                >
                  <span className="relative px-8 py-4 transition-all ease-in duration-75 bg-[#0F111A] rounded-full group-hover:bg-opacity-0 text-xs tracking-[0.2em] font-black uppercase flex items-center gap-2">
                    Iniciar Terminal Elor_TT
                    <ChevronRight className="w-4 h-4 text-zinc-400 group-hover:translate-x-1 transition-transform" />
                  </span>
                </button>

                <div className="text-[9px] text-zinc-500 uppercase tracking-widest font-mono mt-8">
                  Security Code: SEC_NODE_ACTIVE_2026
                </div>
              </div>
            </motion.div>
          ) : !activeDoc ? (
            /* Welcome / Initial Dragzone Landing Stage (Professional Dashboard Launchpad) */
            <motion.div
              key="landing"
              initial={{ opacity: 0, scale: 0.98 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, y: 15 }}
              transition={{ duration: 0.3 }}
              className="flex-1 flex flex-col items-center justify-center my-6 md:my-10"
            >
              
              {/* BRAND CARD HEADER WITH DOUBLE CONTOURS & ROTATED CIRCLE BADGES */}
              <div className="w-full max-w-4xl bg-gradient-to-b from-[#1E2235] to-[#12141F] p-8 md:p-12 rounded-[2.5rem] border-2 border-zinc-700/80 shadow-2xl mb-10 text-center relative overflow-hidden">
                {/* Visual Circle Backdrops */}
                <div className="absolute top-[-30px] left-[-30px] w-24 h-24 rounded-full border border-zinc-800 pointer-events-none" />
                <div className="absolute bottom-[-50px] right-[-50px] w-48 h-48 rounded-full border-2 border-dashed border-zinc-800 pointer-events-none" />
                
                {/* Brand Stamp */}
                <div className="inline-flex items-center gap-2 px-4 py-2 bg-[#0F111A] border-2 border-zinc-700 text-zinc-300 text-xs font-bold uppercase tracking-widest rounded-full mb-6">
                  <span className="w-2.5 h-2.5 rounded-full bg-zinc-500" />
                  Nethels Industries
                </div>

                <h2 className="text-4xl sm:text-5xl font-black text-white tracking-wider mb-8 select-none font-mono">
                  ELOR_TT
                </h2>

                {/* 3 CORE CREATION OPTIONS */}
                <h3 className="text-xs font-bold text-zinc-400 uppercase tracking-[0.2em] mb-4 font-mono select-none">
                  Crear Nuevo Documento o Planilla
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-5 max-w-4xl mx-auto w-full mb-8">
                  {/* Create Empty Sheet */}
                  <div 
                    onClick={handleCreateEmptySpreadsheet}
                    className="group relative bg-[#181C2B] hover:bg-[#202538] p-6 rounded-2xl border-2 border-zinc-700 hover:border-zinc-500 shadow-md transition-all duration-300 cursor-pointer text-left flex flex-col justify-between overflow-hidden"
                  >
                    <div className="absolute top-[-30px] right-[-30px] w-16 h-16 bg-zinc-700/10 rounded-full group-hover:scale-110 transition-transform duration-300" />
                    <div>
                      <div className="w-10 h-10 rounded-full bg-[#272D42] border border-zinc-700 flex items-center justify-center text-white mb-3">
                        <FileSpreadsheet className="w-4.5 h-4.5" />
                      </div>
                      <h4 className="text-sm font-bold text-white mb-1">
                        Nueva Planilla
                      </h4>
                      <div className="text-[10px] text-zinc-400 font-mono font-bold tracking-wider">.xls / .xlsx</div>
                    </div>
                    <div className="mt-4 pt-2 border-t border-zinc-800 text-[10px] font-bold text-zinc-400 uppercase tracking-widest font-mono">
                      Abrir Planilla →
                    </div>
                  </div>

                  {/* Create Empty Word Document */}
                  <div 
                    onClick={handleCreateEmptyDocument}
                    className="group relative bg-[#181C2B] hover:bg-[#202538] p-6 rounded-2xl border-2 border-zinc-700 hover:border-zinc-500 shadow-md transition-all duration-300 cursor-pointer text-left flex flex-col justify-between overflow-hidden"
                  >
                    <div className="absolute top-[-30px] right-[-30px] w-16 h-16 bg-zinc-700/10 rounded-full group-hover:scale-110 transition-transform duration-300" />
                    <div>
                      <div className="w-10 h-10 rounded-full bg-[#272D42] border border-zinc-700 flex items-center justify-center text-white mb-3">
                        <FileText className="w-4.5 h-4.5" />
                      </div>
                      <h4 className="text-sm font-bold text-white mb-1">
                        Nuevo Documento
                      </h4>
                      <div className="text-[10px] text-zinc-400 font-mono font-bold tracking-wider">.doc / .docx</div>
                    </div>
                    <div className="mt-4 pt-2 border-t border-zinc-800 text-[10px] font-bold text-zinc-400 uppercase tracking-widest font-mono">
                      Abrir Editor →
                    </div>
                  </div>

                  {/* Create Empty PDF Document */}
                  <div 
                    onClick={handleCreateEmptyPdf}
                    className="group relative bg-[#181C2B] hover:bg-[#202538] p-6 rounded-2xl border-2 border-zinc-700 hover:border-zinc-500 shadow-md transition-all duration-300 cursor-pointer text-left flex flex-col justify-between overflow-hidden"
                  >
                    <div className="absolute top-[-30px] right-[-30px] w-16 h-16 bg-zinc-700/10 rounded-full group-hover:scale-110 transition-transform duration-300" />
                    <div>
                      <div className="w-10 h-10 rounded-full bg-[#272D42] border border-zinc-700 flex items-center justify-center text-white mb-3">
                        <FileText className="w-4.5 h-4.5" />
                      </div>
                      <h4 className="text-sm font-bold text-white mb-1">
                        Nuevo PDF
                      </h4>
                      <div className="text-[10px] text-zinc-400 font-mono font-bold tracking-wider">.pdf</div>
                    </div>
                    <div className="mt-4 pt-2 border-t border-zinc-800 text-[10px] font-bold text-zinc-400 uppercase tracking-widest font-mono">
                      Abrir PDF →
                    </div>
                  </div>
                </div>
              </div>

              {/* UPLOADER BLOCK WITH CIRCULAR CORNER FRAMEWORK */}
              <div className="w-full max-w-4xl p-2 bg-[#1E2235]/60 border border-zinc-700 rounded-[2rem] shadow-xl">
                <FileUploader onFileUploaded={handleFileUploaded} onError={handleError} />
              </div>

              {/* THREE SPECIFICATION STATS WITH DISTINCT PROFESSIONAL OUTLINES */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 w-full max-w-4xl mt-12 text-zinc-300">
                
                <div className="bg-[#161925] border-2 border-zinc-800 p-6 rounded-2xl shadow-sm hover:border-zinc-700 transition-all text-center relative group">
                  <div className="absolute top-2 right-2 w-2 h-2 rounded-full bg-zinc-600" />
                  <div className="font-bold text-white text-xs uppercase tracking-widest mb-1.5 font-mono">
                    Estructura Elor_TT Grid
                  </div>
                </div>

                <div className="bg-[#161925] border-2 border-zinc-800 p-6 rounded-2xl shadow-sm hover:border-zinc-700 transition-all text-center relative">
                  <div className="absolute top-2 right-2 w-2 h-2 rounded-full bg-zinc-600" />
                  <div className="font-bold text-white text-xs uppercase tracking-widest mb-1.5 font-mono">
                    Mapeador de Texto
                  </div>
                </div>

                <div className="bg-[#161925] border-2 border-zinc-800 p-6 rounded-2xl shadow-sm hover:border-zinc-700 transition-all text-center relative">
                  <div className="absolute top-2 right-2 w-2 h-2 rounded-full bg-zinc-600" />
                  <div className="font-bold text-white text-xs uppercase tracking-widest mb-1.5 font-mono">
                    Nethels Industries Core
                  </div>
                </div>

              </div>
            </motion.div>
          ) : (
            /* Active Analysis Document Hub Full Studio View */
            <motion.div
              key="workspace"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              transition={{ duration: 0.35 }}
              className="flex-1 w-full text-slate-900" /* Force high-contrast text color inside the active document editor */
            >
              <div className="mb-4 flex justify-between items-center bg-[#161925] p-3 rounded-2xl border border-zinc-850 shadow-md">
                <div className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-emerald-500" />
                  <span className="text-xs font-bold text-zinc-300 tracking-wider uppercase">Elor_TT Sesión Activa</span>
                </div>
                <button
                  onClick={handleClearDocument}
                  className="px-3.5 py-1.5 bg-[#2C161B] hover:bg-[#3E1A22] border border-red-900/60 text-red-200 hover:text-white rounded-lg text-[10.5px] font-bold uppercase tracking-wider transition-colors cursor-pointer"
                >
                  Cerrar Editor y Volver al Inicio
                </button>
              </div>

              <DocPreview doc={activeDoc} onClear={handleClearDocument} />
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Modern, clean layout footer */}
      <footer className="bg-[#161925] border-t border-zinc-850 py-6 select-none mt-auto z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-4 text-[10px] text-zinc-400 uppercase tracking-widest font-bold font-mono">
          <div className="flex items-center gap-2 text-zinc-400 font-semibold normal-case tracking-normal">
            <ShieldCheck className="w-4 h-4 text-zinc-400" />
            <span>Suite Elor_TT construida bajo especificación de Nethels Industries</span>
          </div>
          <div className="flex items-center gap-6">
            <span className="hover:text-white transition-colors font-bold text-zinc-300">Elor_TT v2.6.4</span>
            <span>© 2026 Nethels Industries Corporation</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
