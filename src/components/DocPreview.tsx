import React, { useState, useEffect, useMemo, useRef } from "react";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";
import { 
  FileText, 
  FileSpreadsheet, 
  Search, 
  Eye, 
  AlertCircle, 
  RefreshCw, 
  Download, 
  Plus, 
  Trash2, 
  Check, 
  Type, 
  Heading1, 
  Heading2, 
  Heading3,
  Bold, 
  Italic, 
  List, 
  ListOrdered, 
  Underline,
  HelpCircle,
  Save,
  Grid,
  ChevronUp,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  BarChart3,
  Image,
  Upload,
  Indent,
  Outdent,
  Paintbrush,
  Palette,
  Sparkles,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Copy,
  Sliders,
  MousePointer,
  X
} from "lucide-react";
import { ParsedDocument, SheetData } from "../types";

interface DocPreviewProps {
  doc: ParsedDocument;
  onClear: () => void;
}

export default function DocPreview({ doc, onClear }: DocPreviewProps) {
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [editedHtml, setEditedHtml] = useState<string>("");
  const [initialHtml, setInitialHtml] = useState<string>("");
  const [editedText, setEditedText] = useState<string>("");
  
  // Canvas and Selection preservation refs
  const canvasRef = useRef<HTMLDivElement | null>(null);
  const savedRangeRef = useRef<Range | null>(null);

  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  const [spreadsheetSearch, setSpreadsheetSearch] = useState("");
  const [pdfSearch, setPdfSearch] = useState("");

  // Document formatting options
  const [fontFamily, setFontFamily] = useState("font-sans");
  const [fontSize, setFontSize] = useState("15px");
  const [paperBg, setPaperBg] = useState("#FFFFFF");
  const [textColor, setTextColor] = useState("#1F2937");
  const [docHeader, setDocHeader] = useState("");
  const [docFooter, setDocFooter] = useState("");
  const [showPageNumbers, setShowPageNumbers] = useState(true);
  const [showMarginGuides, setShowMarginGuides] = useState(true);
  const [pageNumberFormat, setPageNumberFormat] = useState("pagina-x-de-y");
  const [pageNumberPrefix, setPageNumberPrefix] = useState("Pág.");
  const [pageNumberAlign, setPageNumberAlign] = useState<"left" | "center" | "right">("right");

  // Advanced Header & Footer Customization options
  const [showHeader, setShowHeader] = useState(true);
  const [showFooter, setShowFooter] = useState(true);
  const [headerAlign, setHeaderAlign] = useState<"left" | "center" | "right">("left");
  const [footerAlign, setFooterAlign] = useState<"left" | "center" | "right">("left");
  const [headerBold, setHeaderBold] = useState(true);
  const [headerItalic, setHeaderItalic] = useState(false);
  const [footerBold, setFooterBold] = useState(true);
  const [footerItalic, setFooterItalic] = useState(false);
  const [headerSize, setHeaderSize] = useState("10px");
  const [footerSize, setFooterSize] = useState("10px");
  const [headerColor, setHeaderColor] = useState("#4b5563");
  const [footerColor, setFooterColor] = useState("#4b5563");

  // Math builder states
  const [mathType, setMathType] = useState<"integral" | "raiz" | "sumatoria" | "fraccion" | "potencia">("integral");
  const [mathSize, setMathSize] = useState<"chica" | "mediana" | "grande">("mediana");
  const [integralDefinite, setIntegralDefinite] = useState(true);
  const [integralLower, setIntegralLower] = useState("a");
  const [integralUpper, setIntegralUpper] = useState("b");
  const [integralExpression, setIntegralExpression] = useState("x");

  const [rootType, setRootType] = useState<"square" | "cube" | "custom">("square");
  const [rootIndex, setRootIndex] = useState("n");
  const [rootExpression, setRootExpression] = useState("x");

  const [sumType, setSumType] = useState<"sum" | "prod">("sum");
  const [sumLower, setSumLower] = useState("i=1");
  const [sumUpper, setSumUpper] = useState("n");
  const [sumExpression, setSumExpression] = useState("x");

  const [fracNumerator, setFracNumerator] = useState("x");
  const [fracDenominator, setFracDenominator] = useState("y");
  
  const [powerBase, setPowerBase] = useState("x");
  const [powerExponent, setPowerExponent] = useState("n");
  
  const [activeFormulaInput, setActiveFormulaInput] = useState<string | null>(null);

  const appendSymbolToActiveInput = (symbol: string) => {
    if (!activeFormulaInput) return;
    switch (activeFormulaInput) {
      case "integralLower": setIntegralLower(prev => prev + symbol); break;
      case "integralUpper": setIntegralUpper(prev => prev + symbol); break;
      case "integralExpression": setIntegralExpression(prev => prev + symbol); break;
      case "rootIndex": setRootIndex(prev => prev + symbol); break;
      case "rootExpression": setRootExpression(prev => prev + symbol); break;
      case "sumLower": setSumLower(prev => prev + symbol); break;
      case "sumUpper": setSumUpper(prev => prev + symbol); break;
      case "sumExpression": setSumExpression(prev => prev + symbol); break;
      case "fracNumerator": setFracNumerator(prev => prev + symbol); break;
      case "fracDenominator": setFracDenominator(prev => prev + symbol); break;
      case "powerBase": setPowerBase(prev => prev + symbol); break;
      case "powerExponent": setPowerExponent(prev => prev + symbol); break;
    }
  };

  // Custom Font Uploads state
  const [uploadedFonts, setUploadedFonts] = useState<{ name: string; cssName: string }[]>([]);
  const [isToolbarExpanded, setIsToolbarExpanded] = useState<boolean>(false);

  // New States: Typography Scope, Sizing, Margins & Rulers
  const [typographyScope, setTypographyScope] = useState<"documento" | "seleccion">("seleccion");
  const [globalFontFamily, setGlobalFontFamily] = useState("font-sans");
  const [globalFontSize, setGlobalFontSize] = useState("15px");
  const [globalTextColor, setGlobalTextColor] = useState("#1F2937");

  const [pageSize, setPageSize] = useState<"A4" | "A5" | "Carta" | "Infinita" | "Personalizado">("A4");
  const [customWidth, setCustomWidth] = useState<number>(21.0); // in cm
  const [customHeight, setCustomHeight] = useState<number>(29.7); // in cm

  const [marginLeft, setMarginLeft] = useState<number>(2.5);
  const [marginRight, setMarginRight] = useState<number>(2.5);
  const [marginTop, setMarginTop] = useState<number>(2.5);
  const [marginBottom, setMarginBottom] = useState<number>(2.5);

  const [textIndent, setTextIndent] = useState<number>(0.0);
  const [headerHeight, setHeaderHeight] = useState<number>(1.2);
  const [footerHeight, setFooterHeight] = useState<number>(1.2);

  const [isDragging, setIsDragging] = useState<"marginLeft" | "marginRight" | "marginTop" | "marginBottom" | "indent" | "header" | "footer" | null>(null);
  const [showRulerLines, setShowRulerLines] = useState<boolean>(true);
  const [marginsMode, setMarginsMode] = useState<"2m" | "4m">("4m");

  // Cover page and custom page numbering options
  const [coverHasMargins, setCoverHasMargins] = useState<boolean>(false);
  const [showCoverSelectorModal, setShowCoverSelectorModal] = useState<boolean>(false);
  const [customCoverStyleString, setCustomCoverStyleString] = useState<string>("");
  
  const [pageNumberIncludeCover, setPageNumberIncludeCover] = useState<boolean>(true);
  const [pageNumberStartValue, setPageNumberStartValue] = useState<number>(1);
  const [pageNumberStartAtPageIdx, setPageNumberStartAtPageIdx] = useState<number>(1); // 1-based page index
  const [pageNumberEndAtPageIdx, setPageNumberEndAtPageIdx] = useState<number>(999); // 1-based page index

  const pageDimensions = useMemo(() => {
    switch (pageSize) {
      case "A4":
        return { width: 21.0, height: 29.7 };
      case "A5":
        return { width: 14.8, height: 21.0 };
      case "Carta":
        return { width: 21.6, height: 27.9 };
      case "Infinita":
        return { width: customWidth || 21.0, height: null };
      case "Personalizado":
        return { width: customWidth || 21.0, height: customHeight || 29.7 };
      default:
        return { width: 21.0, height: 29.7 };
    }
  }, [pageSize, customWidth, customHeight]);

  const PAGE_SCALE = 35; // Pixels per centimeter for rulers and paper sizing

  const handleFontFamilyChange = (val: string) => {
    setFontFamily(val);
    if (typographyScope === "documento") {
      setGlobalFontFamily(val);
    } else {
      applyStyleToSelection("font-family", val);
    }
  };

  const handleFontSizeChange = (val: string) => {
    setFontSize(val);
    if (typographyScope === "documento") {
      setGlobalFontSize(val);
    } else {
      applyStyleToSelection("font-size", val);
    }
  };

  const handleTextColorChange = (val: string) => {
    setTextColor(val);
    if (typographyScope === "documento") {
      setGlobalTextColor(val);
    } else {
      applyStyleToSelection("color", val);
    }
  };

  const applyIndentToSelection = (indentCm: number) => {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const range = selection.getRangeAt(0);
    
    let element = range.startContainer.nodeType === Node.ELEMENT_NODE 
      ? (range.startContainer as HTMLElement) 
      : range.startContainer.parentElement;
      
    while (element && element.id !== "rich-doc-canvas" && !["P", "DIV", "H1", "H2", "H3", "LI"].includes(element.tagName)) {
      element = element.parentElement;
    }
    
    if (element && element.id !== "rich-doc-canvas") {
      element.style.textIndent = indentCm !== 0 ? `${indentCm}cm` : "";
      const canvas = document.getElementById("rich-doc-canvas") || canvasRef.current;
      if (canvas) {
        setEditedHtml(canvas.innerHTML);
      }
    }
  };

  const handleStartDrag = (
    e: React.MouseEvent,
    type: "marginLeft" | "marginRight" | "marginTop" | "marginBottom" | "indent" | "header" | "footer",
    containerSizeCm: number,
    containerSizePx: number
  ) => {
    e.preventDefault();
    const startX = e.clientX;
    const startY = e.clientY;
    
    let initialValue = 0;
    if (type === "marginLeft") initialValue = marginLeft;
    else if (type === "marginRight") initialValue = marginRight;
    else if (type === "marginTop") initialValue = marginTop;
    else if (type === "marginBottom") initialValue = marginBottom;
    else if (type === "indent") initialValue = textIndent;
    else if (type === "header") initialValue = headerHeight;
    else if (type === "footer") initialValue = footerHeight;

    setIsDragging(type);

    const onMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.clientX - startX;
      const deltaY = moveEvent.clientY - startY;
      const pxToCm = containerSizeCm / containerSizePx;
      
      if (type === "marginLeft" || type === "marginRight" || type === "indent") {
        const deltaCm = deltaX * pxToCm;
        if (type === "marginLeft") {
          const newVal = Math.max(0.5, Math.min(containerSizeCm - marginRight - 1.0, initialValue + deltaCm));
          setMarginLeft(parseFloat(newVal.toFixed(1)));
        } else if (type === "marginRight") {
          const newVal = Math.max(0.5, Math.min(containerSizeCm - marginLeft - 1.0, initialValue - deltaCm));
          setMarginRight(parseFloat(newVal.toFixed(1)));
        } else if (type === "indent") {
          const newVal = Math.max(-marginLeft + 0.5, Math.min(10.0, initialValue + deltaCm));
          setTextIndent(parseFloat(newVal.toFixed(1)));
          applyIndentToSelection(parseFloat(newVal.toFixed(1)));
        }
      } else {
        const deltaCm = deltaY * pxToCm;
        if (type === "marginTop") {
          const newVal = Math.max(0.5, Math.min(containerSizeCm - marginBottom - 1.0, initialValue + deltaCm));
          setMarginTop(parseFloat(newVal.toFixed(1)));
        } else if (type === "marginBottom") {
          const newVal = Math.max(0.5, Math.min(containerSizeCm - marginTop - 1.0, initialValue - deltaCm));
          setMarginBottom(parseFloat(newVal.toFixed(1)));
        } else if (type === "header") {
          const newVal = Math.max(0.1, Math.min(marginTop - 0.2, initialValue + deltaCm));
          setHeaderHeight(parseFloat(newVal.toFixed(1)));
        } else if (type === "footer") {
          const newVal = Math.max(0.1, Math.min(marginBottom - 0.2, initialValue - deltaCm));
          setFooterHeight(parseFloat(newVal.toFixed(1)));
        }
      }
    };

    const onMouseUp = () => {
      setIsDragging(null);
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
    };

    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
  };

  // Collapsible toolbar sections state
  const [openSections, setOpenSections] = useState<{
    inicio: boolean;
    insertar: boolean;
    pagina: boolean;
    simbolos: boolean;
    configPag: boolean;
    ajustes: boolean;
  }>({
    inicio: true,
    insertar: false,
    pagina: false,
    simbolos: false,
    configPag: false,
    ajustes: false
  });

  // Page navigation mode state: scroll (Continuous scroll) vs pages (Individual pages with arrows)
  const [navigationMode, setNavigationMode] = useState<"scroll" | "pages">("scroll");

  // Page management states
  const [docPages, setDocPages] = useState<string[]>([""]);
  const [activePageIdx, setActivePageIdx] = useState<number>(0);
  const [selectedPages, setSelectedPages] = useState<number[]>([]);

  const scrollToPage = (idx: number) => {
    if (navigationMode === "scroll") {
      const pageEl = document.getElementById(`paper-sheet-container-${idx}`);
      if (pageEl) {
        pageEl.scrollIntoView({ behavior: "smooth", block: "center" });
      }
    }
  };

  const goToNextPage = () => {
    if (activePageIdx < docPages.length - 1) {
      setActivePageIdx(prev => {
        const next = prev + 1;
        setTimeout(() => scrollToPage(next), 50);
        return next;
      });
    } else {
      handleAddNewPage();
      setTimeout(() => scrollToPage(docPages.length), 150);
    }
  };

  const goToPrevPage = () => {
    if (activePageIdx > 0) {
      setActivePageIdx(prev => {
        const prevIdx = prev - 1;
        setTimeout(() => scrollToPage(prevIdx), 50);
        return prevIdx;
      });
    }
  };

  const stats = useMemo(() => {
    // Active page text
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = editedHtml || docPages[activePageIdx] || "";
    const activeText = tempDiv.innerText || tempDiv.textContent || "";

    // All text
    const allText = docPages.map((page, idx) => {
      if (idx === activePageIdx) {
        return activeText;
      }
      const pDiv = document.createElement("div");
      pDiv.innerHTML = page || "";
      return pDiv.innerText || pDiv.textContent || "";
    }).join(" ");

    const getWordCount = (txt: string) => {
      const words = txt.trim().split(/\s+/).filter(w => w.length > 0);
      return words.length;
    };

    return {
      activeWords: getWordCount(activeText),
      activeChars: activeText.length,
      totalWords: getWordCount(allText),
      totalChars: allText.length,
    };
  }, [editedHtml, docPages, activePageIdx]);

  const lastPageAddRef = useRef<number>(0);

  const handleAddNewPage = () => {
    if (pageSize === "Infinita") return;
    
    let currentHtml = "<p><br></p>";
    if (canvasRef.current) {
      currentHtml = canvasRef.current.innerHTML;
    }
    
    setDocPages(prev => {
      const updated = [...prev];
      updated[activePageIdx] = currentHtml;
      return [...updated, "<p><br></p>"];
    });
    
    setActivePageIdx(docPages.length);
  };

  const handleWheelAddPage = (e: React.WheelEvent<HTMLDivElement>) => {
    if (pageSize === "Infinita") return;
    if (activePageIdx !== docPages.length - 1) return; // Only at the last page

    // Check if scrolling down
    if (e.deltaY > 60) {
      const container = e.currentTarget;
      const isAtBottom = Math.abs(container.scrollHeight - container.clientHeight - container.scrollTop) < 15;
      
      if (isAtBottom) {
        const now = Date.now();
        if (now - lastPageAddRef.current > 1500) { // 1.5 seconds cooldown
          lastPageAddRef.current = now;
          handleAddNewPage();
        }
      }
    }
  };

  const handleKeyDownAddPage = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (pageSize === "Infinita") return;
    if (activePageIdx !== docPages.length - 1) return; // Only at the last page

    if (e.key === "ArrowDown" || e.key === "ArrowRight") {
      // Check if cursor is at the end of the text
      const selection = window.getSelection();
      if (selection && selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        const canvas = canvasRef.current;
        if (canvas) {
          // Check if caret is at the last block/character
          const preCaretRange = range.cloneRange();
          preCaretRange.selectNodeContents(canvas);
          preCaretRange.setStart(range.endContainer, range.endOffset);
          const textAfter = preCaretRange.toString().trim();
          
          if (textAfter.length === 0) {
            e.preventDefault();
            handleAddNewPage();
          }
        }
      }
    }
  };

  const handleWorksheetScroll = (e: React.UIEvent<HTMLDivElement>) => {
    if (navigationMode !== "scroll") return;
    const container = e.currentTarget;
    const containerCenter = container.getBoundingClientRect().top + container.clientHeight / 2;
    
    let closestPageIdx = activePageIdx;
    let minDistance = Infinity;
    
    for (let i = 0; i < docPages.length; i++) {
      const pageEl = document.getElementById(`paper-sheet-container-${i}`);
      if (pageEl) {
        const rect = pageEl.getBoundingClientRect();
        const pageCenter = rect.top + rect.height / 2;
        const distance = Math.abs(pageCenter - containerCenter);
        if (distance < minDistance) {
          minDistance = distance;
          closestPageIdx = i;
        }
      }
    }
    
    if (closestPageIdx !== activePageIdx) {
      setActivePageIdx(closestPageIdx);
    }
  };

  // Copied text formatting state
  const [copiedFormat, setCopiedFormat] = useState<{
    color?: string;
    fontFamily?: string;
    fontSize?: string;
    fontWeight?: string;
    fontStyle?: string;
    textDecoration?: string;
  } | null>(null);

  const toggleSection = (section: "inicio" | "insertar" | "pagina" | "simbolos" | "configPag" | "ajustes") => {
    setOpenSections(prev => {
      const isAlreadyOpen = prev[section as keyof typeof prev];
      return {
        inicio: section === "inicio" ? !isAlreadyOpen : false,
        insertar: section === "insertar" ? !isAlreadyOpen : false,
        pagina: section === "pagina" ? !isAlreadyOpen : false,
        simbolos: section === "simbolos" ? !isAlreadyOpen : false,
        configPag: section === "configPag" ? !isAlreadyOpen : false,
        ajustes: section === "ajustes" ? !isAlreadyOpen : false,
      };
    });
  };

  const handleCopyFormat = () => {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const range = selection.getRangeAt(0);
    let element = range.startContainer as HTMLElement;
    if (element.nodeType === Node.TEXT_NODE) {
      element = element.parentElement as HTMLElement;
    }
    
    if (element) {
      const computed = window.getComputedStyle(element);
      setCopiedFormat({
        color: computed.color,
        fontFamily: computed.fontFamily,
        fontSize: computed.fontSize,
        fontWeight: computed.fontWeight,
        fontStyle: computed.fontStyle,
        textDecoration: computed.textDecoration
      });
    }
  };

  const handlePasteFormat = () => {
    if (!copiedFormat) return;
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const range = selection.getRangeAt(0);

    const span = document.createElement("span");
    if (copiedFormat.color) span.style.color = copiedFormat.color;
    if (copiedFormat.fontFamily) span.style.fontFamily = copiedFormat.fontFamily;
    if (copiedFormat.fontSize) span.style.fontSize = copiedFormat.fontSize;
    if (copiedFormat.fontWeight) span.style.fontWeight = copiedFormat.fontWeight;
    if (copiedFormat.fontStyle) span.style.fontStyle = copiedFormat.fontStyle;
    if (copiedFormat.textDecoration) span.style.textDecoration = copiedFormat.textDecoration;

    try {
      range.surroundContents(span);
    } catch (e) {
      // Fallback
      if (copiedFormat.color) {
        document.execCommand("styleWithCSS", false, "true");
        document.execCommand("foreColor", false, copiedFormat.color);
      }
    }

    const canvas = document.getElementById("rich-doc-canvas");
    if (canvas) {
      setEditedHtml(canvas.innerHTML);
    }
  };

  // Spreadsheet Live Chart options
  const [showSpreadsheetChart, setShowSpreadsheetChart] = useState(false);
  const [chartType, setChartType] = useState<"bar" | "line" | "area" | "pie">("bar");
  const [xAxisCol, setXAxisCol] = useState(0);
  const [yAxisCol, setYAxisCol] = useState(1);
  const [chartTitle, setChartTitle] = useState("Análisis de Tendencia Elor_TT");

  // Sheet creation states
  const [isCreatingSheet, setIsCreatingSheet] = useState(false);
  const [newSheetName, setNewSheetName] = useState("");

  // Infinite display boundaries
  const [visibleRowsCount, setVisibleRowsCount] = useState(100);
  const [visibleColsCount, setVisibleColsCount] = useState(26);

  // Cell Editing coordinates State
  const [editingCell, setEditingCell] = useState<{ rowIdx: number; colIdx: number } | null>(null);
  const [editValue, setEditValue] = useState<string>("");

  // Search Navigation coordinates
  const [currentMatchIdx, setCurrentMatchIdx] = useState(0);

  // Load Google Fonts for premium workspace look dynamically
  useEffect(() => {
    const link = document.createElement("link");
    link.href = "https://fonts.googleapis.com/css2?family=Cinzel:wght@600;800&family=Merriweather:ital,wght@0,400;0,700;1,400&family=Space+Grotesk:wght@500;700&family=Playfair+Display:ital,wght@0,700;1,400&family=JetBrains+Mono:wght@400;700&display=swap";
    link.rel = "stylesheet";
    document.head.appendChild(link);
    return () => {
      document.head.removeChild(link);
    };
  }, []);

  // Helper to split a document's HTML by page break elements
  const parsePagesFromHtml = (html: string): string[] => {
    if (!html) return ["<p><br></p>"];
    const separatorRegex = /<hr[^>]*page-break-after:\s*always;[^>]*>|<div[^>]*page-break-after:\s*always;[^>]*><\/div>|<!--\s*page-break\s*-->/gi;
    const parts = html.split(separatorRegex);
    const cleaned = parts.map(p => p.trim()).filter(Boolean);
    return cleaned.length > 0 ? cleaned : [html];
  };

  const movePageUp = (idx: number) => {
    if (idx <= 0) return;
    setDocPages(prev => {
      const updated = [...prev];
      const temp = updated[idx];
      updated[idx] = updated[idx - 1];
      updated[idx - 1] = temp;
      return updated;
    });
    if (activePageIdx === idx) {
      setActivePageIdx(idx - 1);
    } else if (activePageIdx === idx - 1) {
      setActivePageIdx(idx);
    }
  };

  const movePageDown = (idx: number) => {
    if (idx >= docPages.length - 1) return;
    setDocPages(prev => {
      const updated = [...prev];
      const temp = updated[idx];
      updated[idx] = updated[idx + 1];
      updated[idx + 1] = temp;
      return updated;
    });
    if (activePageIdx === idx) {
      setActivePageIdx(idx + 1);
    } else if (activePageIdx === idx + 1) {
      setActivePageIdx(idx);
    }
  };

  const moveSelectedPagesUp = () => {
    if (selectedPages.length === 0) return;
    const sortedIndices = [...selectedPages].sort((a, b) => a - b);
    if (sortedIndices[0] === 0) return;
    
    setDocPages(prev => {
      const updated = [...prev];
      for (const idx of sortedIndices) {
        const temp = updated[idx];
        updated[idx] = updated[idx - 1];
        updated[idx - 1] = temp;
      }
      return updated;
    });
    if (selectedPages.includes(activePageIdx) && activePageIdx > 0) {
      setActivePageIdx(activePageIdx - 1);
    }
    setSelectedPages(prev => prev.map(idx => idx > 0 ? idx - 1 : idx));
  };

  const deleteSelectedPages = () => {
    if (selectedPages.length === 0) return;
    if (selectedPages.length === docPages.length) {
      alert("No puedes eliminar todas las páginas del documento.");
      return;
    }
    
    setDocPages(prev => {
      return prev.filter((_, idx) => !selectedPages.includes(idx));
    });
    
    let newActiveIdx = activePageIdx;
    while (newActiveIdx >= docPages.length - selectedPages.length) {
      newActiveIdx--;
    }
    if (newActiveIdx < 0) newActiveIdx = 0;
    setActivePageIdx(newActiveIdx);
    setSelectedPages([]);
  };

  // Populate local editor state when document changes
  useEffect(() => {
    if (doc.visualContent?.sheets) {
      setSheets(doc.visualContent.sheets);
    } else {
      // Create at least one sheet if empty
      setSheets([{
        sheetName: "Hoja 1",
        grid: Array.from({ length: 40 }).map(() => Array.from({ length: 15 }).map(() => "")),
        rowCount: 40,
        columnCount: 15
      }]);
    }
    
    let htmlContent = "";
    if (doc.visualContent?.html) {
      htmlContent = doc.visualContent.html;
    } else if (doc.extension.toLowerCase() === ".pdf") {
      // PDF text extracted becomes beautiful editable blocks
      const rawText = doc.visualContent?.excerpt || doc.extractedText || "";
      htmlContent = rawText
        ? rawText
            .split("\n")
            .filter(line => line.trim() !== "")
            .map(line => `<p>${line.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</p>`)
            .join("")
        : "";
    }
    
    const parsed = parsePagesFromHtml(htmlContent);
    setDocPages(parsed);
    setActivePageIdx(0);
    setSelectedPages([]);
    
    const firstPageHtml = parsed[0] || "<p><br></p>";
    setInitialHtml(firstPageHtml);
    setEditedHtml(htmlContent || firstPageHtml);

    setEditedText(doc.visualContent?.excerpt || doc.extractedText || "");
    setActiveSheetIdx(0);
    setSpreadsheetSearch("");
    setPdfSearch("");
    setEditingCell(null);
    setCurrentMatchIdx(0);
  }, [doc]);

  // Populate contentEditable canvas safely only once when new initialHtml is loaded or when canvas initializes
  useEffect(() => {
    if (canvasRef.current) {
      canvasRef.current.innerHTML = initialHtml;
    }
  }, [initialHtml]);

  const lastSyncedPageIdxRef = useRef<number>(activePageIdx);

  // Load current page HTML into canvas when activePageIdx changes
  useEffect(() => {
    if (docPages[activePageIdx] !== undefined) {
      setEditedHtml(docPages[activePageIdx]);
      if (canvasRef.current && canvasRef.current.innerHTML !== docPages[activePageIdx]) {
        canvasRef.current.innerHTML = docPages[activePageIdx];
      }
    }
  }, [activePageIdx, docPages]);

  // Synchronize canvas edits to docPages array
  useEffect(() => {
    if (lastSyncedPageIdxRef.current !== activePageIdx) {
      lastSyncedPageIdxRef.current = activePageIdx;
      return;
    }
    if (docPages[activePageIdx] !== undefined && editedHtml !== docPages[activePageIdx]) {
      setDocPages(prev => {
        const updated = [...prev];
        updated[activePageIdx] = editedHtml;
        return updated;
      });
    }
  }, [editedHtml, activePageIdx, docPages]);

  // Read active sheet accurately
  const activeSheet = sheets[activeSheetIdx] || sheets[0];

  // Dynamically configure snug dimensions based on active sheet's actual content to keep scroll bars big and usable
  useEffect(() => {
    if (activeSheet) {
      let maxRowWithData = 0;
      let maxColWithData = 0;
      activeSheet.grid.forEach((row, rIdx) => {
        row.forEach((cell, cIdx) => {
          if (cell !== undefined && cell !== null && String(cell).trim() !== "") {
            maxRowWithData = Math.max(maxRowWithData, rIdx);
            maxColWithData = Math.max(maxColWithData, cIdx);
          }
        });
      });
      // Snug initial layout based on maximum filled data row + 15 rows (default min 30 rows)
      setVisibleRowsCount(Math.max(maxRowWithData + 15, 30));
      // Snug initial layout based on maximum filled data col + 4 cols (default min 11 cols)
      setVisibleColsCount(Math.max(maxColWithData + 4, 11));
      setCurrentMatchIdx(0);
    }
  }, [activeSheetIdx, activeSheet]);

  const formatFileSize = (bytes: number) => {
    if (!bytes || bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
  };

  // Convert column index (0, 1, 2...) to Excel column letters
  const getColLetter = (index: number): string => {
    let letter = "";
    let temp = index;
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  };

  // Extract list of all column letter options for spreadsheet configurations
  const allColumnsList = useMemo(() => {
    if (!activeSheet) return [];
    return Array.from({ length: Math.min(activeSheet.columnCount, 26) }).map((_, idx) => ({
      colIdx: idx,
      letter: getColLetter(idx)
    }));
  }, [activeSheet]);

  // Extract column indexes that hold numerical values
  const numericalColumnsList = useMemo(() => {
    if (!activeSheet) return [];
    const cols: { colIdx: number; letter: string }[] = [];
    const maxColsToInspect = Math.min(activeSheet.columnCount, 15);
    
    for (let c = 0; c < maxColsToInspect; c++) {
      let numericCount = 0;
      let nonNumericCount = 0;
      const rowsToInspect = Math.min(activeSheet.grid.length, 25);
      for (let r = 1; r < rowsToInspect; r++) {
        const val = activeSheet.grid[r]?.[c];
        if (val !== undefined && val !== null && val !== "") {
          if (!isNaN(Number(val))) {
            numericCount++;
          } else {
            nonNumericCount++;
          }
        }
      }
      if (numericCount > 0) {
        cols.push({ colIdx: c, letter: getColLetter(c) });
      }
    }
    return cols;
  }, [activeSheet]);

  // Extract chart points from the active grid dynamically
  const spreadsheetChartData = useMemo(() => {
    if (!activeSheet) return [];
    const rows = activeSheet.grid;
    const result: { label: string; val: number; rIdx: number }[] = [];
    
    const startRow = 1; // Row 0 is headers
    const maxScanRows = Math.min(rows.length, 45);
    
    for (let r = startRow; r < maxScanRows; r++) {
      const row = rows[r];
      if (!row) continue;
      
      const labelVal = row[xAxisCol];
      const numericVal = row[yAxisCol];
      
      const labelStr = labelVal !== undefined && labelVal !== null && String(labelVal).trim() !== "" 
        ? String(labelVal) 
        : `Fila ${r + 1}`;
        
      const num = Number(numericVal);
      if (!isNaN(num) && numericVal !== "" && numericVal !== undefined && numericVal !== null) {
        result.push({
          label: labelStr,
          val: num,
          rIdx: r
        });
      }
    }
    return result;
  }, [activeSheet, xAxisCol, yAxisCol]);

  // Find all match coordinates inside the full grid of activeSheet in real-time
  const searchMatches = useMemo(() => {
    const result: { rowIdx: number; colIdx: number }[] = [];
    const term = spreadsheetSearch.toLowerCase().trim();
    if (!term || !activeSheet) return result;

    for (let r = 0; r < activeSheet.grid.length; r++) {
      const row = activeSheet.grid[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const val = String(row[c] ?? "").toLowerCase();
        if (val.includes(term)) {
          result.push({ rowIdx: r, colIdx: c });
        }
      }
    }
    return result;
  }, [activeSheet, spreadsheetSearch]);

  // Adjust match index bounds if matches list changes
  useEffect(() => {
    setCurrentMatchIdx(0);
  }, [searchMatches.length]);

  // Smooth scroll active match cell into precise visual view inside scroll container
  useEffect(() => {
    if (searchMatches.length > 0 && searchMatches[currentMatchIdx]) {
      const activeMatchObj = searchMatches[currentMatchIdx];
      
      // Auto expand infinite visual view boundaries if target match is beyond current bounds
      let boundaryChanged = false;
      if (activeMatchObj.rowIdx >= visibleRowsCount) {
        setVisibleRowsCount(activeMatchObj.rowIdx + 15);
        boundaryChanged = true;
      }
      if (activeMatchObj.colIdx >= visibleColsCount) {
        setVisibleColsCount(activeMatchObj.colIdx + 5);
        boundaryChanged = true;
      }
      
      const scrollToCell = () => {
        const cellId = `cell-${activeMatchObj.rowIdx}-${activeMatchObj.colIdx}`;
        const elem = document.getElementById(cellId);
        if (elem) {
          elem.scrollIntoView({
            behavior: "smooth",
            block: "center",
            inline: "center"
          });
        }
      };

      if (boundaryChanged) {
        const timer = setTimeout(scrollToCell, 80);
        return () => clearTimeout(timer);
      } else {
        scrollToCell();
      }
    }
  }, [currentMatchIdx, searchMatches, visibleRowsCount, visibleColsCount]);

  // Handle cell changing value
  const handleStartEditing = (rowIdx: number, colIdx: number, currentVal: any) => {
    setEditingCell({ rowIdx, colIdx });
    setEditValue(currentVal !== undefined && currentVal !== null ? String(currentVal) : "");
  };

  const handleFinishEditing = (rowIdx: number, colIdx: number) => {
    if (!activeSheet) return;
    
    const updatedSheets = sheets.map((sh, sIdx) => {
      if (sIdx !== activeSheetIdx) return sh;
      
      // Deep clone current grid to safely make surgical cells insertion
      const updatedGrid = sh.grid.map(row => [...row]);
      
      // Pad empty rows if we edited a cell outside boundaries
      while (updatedGrid.length <= rowIdx) {
        updatedGrid.push(Array(sh.columnCount).fill(""));
      }
      
      // Ensure all rows have enough column spots
      const nextColCount = Math.max(sh.columnCount, colIdx + 1);
      updatedGrid.forEach((row) => {
        while (row.length < nextColCount) {
          row.push("");
        }
      });
      
      const valStr = editValue.trim();
      const parsedVal = !isNaN(Number(valStr)) && valStr !== "" ? Number(valStr) : editValue;
      
      updatedGrid[rowIdx][colIdx] = parsedVal;
      
      return {
        ...sh,
        grid: updatedGrid,
        rowCount: updatedGrid.length,
        columnCount: nextColCount
      };
    });
    
    setSheets(updatedSheets);
    setEditingCell(null);

    // Expand snug layout boundaries dynamically if they edit cells close to the margins
    if (rowIdx >= visibleRowsCount - 4) {
      setVisibleRowsCount(prev => Math.max(prev, rowIdx + 15));
    }
    if (colIdx >= visibleColsCount - 2) {
      setVisibleColsCount(prev => Math.max(prev, colIdx + 4));
    }
  };

  // Add Row to current grid
  const handleAddRow = () => {
    if (!activeSheet) return;
    
    const updatedSheets = sheets.map((sh, sIdx) => {
      if (sIdx !== activeSheetIdx) return sh;
      
      const newRow = Array.from({ length: sh.columnCount }).map(() => "");
      const updatedGrid = [...sh.grid, newRow];
      
      return {
        ...sh,
        grid: updatedGrid,
        rowCount: updatedGrid.length
      };
    });
    
    setSheets(updatedSheets);
    setVisibleRowsCount(prev => Math.max(prev, activeSheet.grid.length + 10));
  };

  // Add Column to current grid
  const handleAddCol = () => {
    if (!activeSheet) return;
    
    const updatedSheets = sheets.map((sh, sIdx) => {
      if (sIdx !== activeSheetIdx) return sh;
      
      const updatedGrid = sh.grid.map(row => [...row, ""]);
      const newColumnVal = sh.columnCount + 1;
      
      return {
        ...sh,
        grid: updatedGrid,
        columnCount: newColumnVal
      };
    });
    
    setSheets(updatedSheets);
    setVisibleColsCount(prev => Math.max(prev, activeSheet.columnCount + 3));
  };

  // Create clean new Sheet tab instantly
  const handleCreateSheet = (e: React.FormEvent) => {
    e.preventDefault();
    const cleanName = newSheetName.trim() || `Hoja ${sheets.length + 1}`;
    
    if (sheets.some(s => s.sheetName.toLowerCase() === cleanName.toLowerCase())) {
      alert("Ya existe una hoja con ese nombre.");
      return;
    }

    // Set up standard 40x15 initial default empty coordinates
    const initialGrid = Array.from({ length: 40 }).map(() => 
      Array.from({ length: 15 }).map(() => "")
    );

    const newSheet: SheetData = {
      sheetName: cleanName,
      grid: initialGrid,
      rowCount: 40,
      columnCount: 15
    };

    const updated = [...sheets, newSheet];
    setSheets(updated);
    setActiveSheetIdx(updated.length - 1);
    setNewSheetName("");
    setIsCreatingSheet(false);
  };

  // Render a beautifully formatted responsive live SVG chart based on current sheet data
  const renderSpreadsheetLiveChart = () => {
    if (spreadsheetChartData.length === 0) {
      return (
        <div className="flex flex-col items-center justify-center p-8 text-center text-zinc-400 bg-zinc-50 border border-dashed border-zinc-200 rounded-2xl h-72">
          <BarChart3 className="w-10 h-10 text-zinc-300 mb-2.5 animate-pulse" />
          <p className="text-xs font-mono font-bold uppercase tracking-wider text-zinc-700">Sin datos numéricos válidos</p>
          <p className="text-[10px] text-zinc-400 mt-1.5 max-w-xs leading-relaxed">
            Ingresa números en cualquier columna secundaria (ej. columna B, C) y etiquetas de texto en la columna A para trazar gráficos dinámicos de forma instantánea.
          </p>
        </div>
      );
    }

    const svgWidth = 500;
    const svgHeight = 260;
    const paddingLeft = 50;
    const paddingRight = 30;
    const paddingTop = 40;
    const paddingBottom = 50;
    
    const chartWidth = svgWidth - paddingLeft - paddingRight;
    const chartHeight = svgHeight - paddingTop - paddingBottom;
    
    const values = spreadsheetChartData.map(d => d.val);
    const maxVal = Math.max(...values, 0) * 1.18 || 10;
    const minVal = 0;

    const stepX = chartWidth / Math.max(spreadsheetChartData.length - 1, 1);
    const barWidth = Math.max(6, Math.min(38, (chartWidth / spreadsheetChartData.length) * 0.6));

    const colors = ["#18181b", "#3f3f46", "#71717a", "#a1a1aa", "#d4d4d8"];

    return (
      <div className="bg-zinc-50 border border-zinc-200 p-5 rounded-2xl flex flex-col gap-4">
        {/* Chart Customizations panel */}
        <div className="grid grid-cols-2 gap-2.5 bg-white border border-zinc-200 p-3 rounded-xl select-none">
          <div>
            <label className="block text-[9px] font-bold text-zinc-400 uppercase tracking-wider font-mono mb-1">Título del Gráfico</label>
            <input
              type="text"
              value={chartTitle}
              onChange={(e) => setChartTitle(e.target.value)}
              className="w-full bg-zinc-50 border border-zinc-200 text-[11px] rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-black font-sans font-medium"
            />
          </div>
          <div>
            <label className="block text-[9px] font-bold text-zinc-400 uppercase tracking-wider font-mono mb-1">Tipo de Gráfico</label>
            <select
              value={chartType}
              onChange={(e) => setChartType(e.target.value as any)}
              className="w-full bg-zinc-50 border border-zinc-200 text-[11px] rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-black font-sans font-medium"
            >
              <option value="bar">📊 Columnas (Barras)</option>
              <option value="line">📈 Línea de Tendencia</option>
              <option value="area">📉 Área Sombreada</option>
              <option value="pie">🍕 Circular (Dona)</option>
            </select>
          </div>
          <div>
            <label className="block text-[9px] font-bold text-zinc-400 uppercase tracking-wider font-mono mb-1">Eje X (Nombres)</label>
            <select
              value={xAxisCol}
              onChange={(e) => setXAxisCol(Number(e.target.value))}
              className="w-full bg-zinc-50 border border-zinc-200 text-[11px] rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-black font-sans font-medium"
            >
              {allColumnsList.map(c => (
                <option key={c.colIdx} value={c.colIdx}>Columna {c.letter}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-[9px] font-bold text-zinc-400 uppercase tracking-wider font-mono mb-1">Eje Y (Valores)</label>
            <select
              value={yAxisCol}
              onChange={(e) => setYAxisCol(Number(e.target.value))}
              className="w-full bg-zinc-50 border border-zinc-200 text-[11px] rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-black font-sans font-medium"
            >
              {allColumnsList.map(c => (
                <option key={c.colIdx} value={c.colIdx}>Columna {c.letter}</option>
              ))}
            </select>
          </div>
        </div>

        {/* Dynamic SVG Drawing */}
        {chartType === "pie" ? (
          <div className="flex flex-col sm:flex-row items-center justify-center gap-6 bg-white border border-zinc-200 p-4 rounded-xl">
            <div className="relative w-full max-w-[160px] shrink-0">
              <svg viewBox="0 0 200 200" className="w-full h-auto">
                {(() => {
                  const total = values.reduce((sum, v) => sum + v, 0);
                  let cumulativeAngle = 0;
                  return spreadsheetChartData.map((d, idx) => {
                    const percentage = total > 0 ? d.val / total : 0;
                    const angle = percentage * 360;
                    
                    const x1 = 100 + 80 * Math.cos((cumulativeAngle - 90) * Math.PI / 180);
                    const y1 = 100 + 80 * Math.sin((cumulativeAngle - 90) * Math.PI / 180);
                    cumulativeAngle += angle;
                    const x2 = 100 + 80 * Math.cos((cumulativeAngle - 90) * Math.PI / 180);
                    const y2 = 100 + 80 * Math.sin((cumulativeAngle - 90) * Math.PI / 180);
                    
                    const largeArcFlag = angle > 180 ? 1 : 0;
                    const color = colors[idx % colors.length];
                    
                    return (
                      <path
                        key={idx}
                        d={`M 100 100 L ${x1} ${y1} A 80 80 0 ${largeArcFlag} 1 ${x2} ${y2} Z`}
                        fill={color}
                        stroke="#ffffff"
                        strokeWidth="1.5"
                        className="hover:opacity-85 transition-opacity"
                      >
                        <title>{d.label}: {d.val}</title>
                      </path>
                    );
                  });
                })()}
                <circle cx="100" cy="100" r="45" fill="#ffffff" />
              </svg>
            </div>
            <div className="flex-1 text-left">
              <h5 className="text-[10px] font-bold text-zinc-950 uppercase tracking-widest font-mono mb-2">{chartTitle}</h5>
              <div className="space-y-1.5 max-h-40 overflow-y-auto pr-1">
                {(() => {
                  const total = values.reduce((sum, v) => sum + v, 0);
                  return spreadsheetChartData.map((d, idx) => {
                    const percentage = total > 0 ? d.val / total : 0;
                    return (
                      <div key={idx} className="flex items-center justify-between text-[10px] text-zinc-600 font-mono">
                        <div className="flex items-center gap-1.5 min-w-0">
                          <span className="w-2 h-2 rounded shrink-0" style={{ backgroundColor: colors[idx % colors.length] }} />
                          <span className="truncate font-semibold text-zinc-800">{d.label}</span>
                        </div>
                        <span className="font-bold text-zinc-500 shrink-0 ml-2">{d.val} ({Math.round(percentage * 100)}%)</span>
                      </div>
                    );
                  });
                })()}
              </div>
            </div>
          </div>
        ) : (
          <div className="bg-white border border-zinc-200 p-3 rounded-xl">
            <h5 className="text-[10px] font-bold text-zinc-950 uppercase tracking-widest font-mono text-left mb-2 pl-2">{chartTitle}</h5>
            <svg viewBox={`0 0 ${svgWidth} ${svgHeight}`} className="w-full h-auto">
              {/* Grid Y Guidelines */}
              {Array.from({ length: 5 }).map((_, i) => {
                const yVal = paddingTop + (chartHeight / 4) * i;
                const displayLabel = Math.round(maxVal - (maxVal / 4) * i);
                return (
                  <g key={i} className="opacity-45">
                    <line x1={paddingLeft} y1={yVal} x2={svgWidth - paddingRight} y2={yVal} stroke="#e4e4e7" strokeDasharray="3,3" strokeWidth="1" />
                    <text x={paddingLeft - 8} y={yVal + 3} fontSize="8" fontFamily="monospace" textAnchor="end" fill="#71717a">
                      {displayLabel}
                    </text>
                  </g>
                );
              })}
              
              {/* Dynamic plotting elements */}
              {spreadsheetChartData.map((d, idx) => {
                const x = paddingLeft + (spreadsheetChartData.length === 1 ? chartWidth / 2 : stepX * idx);
                const y = svgHeight - paddingBottom - ((d.val - minVal) / (maxVal - minVal)) * chartHeight;
                const height = svgHeight - paddingBottom - y;
                
                return (
                  <g key={idx} className="group">
                    {chartType === "bar" && (
                      <rect
                        x={x - barWidth / 2}
                        y={y}
                        width={barWidth}
                        height={Math.max(2, height)}
                        fill="#18181b"
                        rx="3"
                        className="hover:fill-zinc-600 transition-colors"
                      >
                        <title>{d.label}: {d.val}</title>
                      </rect>
                    )}
                    
                    {(chartType === "line" || chartType === "area") && (
                      <circle
                        cx={x}
                        cy={y}
                        r="4"
                        fill="#18181b"
                        stroke="#ffffff"
                        strokeWidth="1.5"
                        className="hover:scale-125 transition-transform"
                      >
                        <title>{d.label}: {d.val}</title>
                      </circle>
                    )}
                    
                    <text
                      x={x}
                      y={svgHeight - paddingBottom + 14}
                      fontSize="7"
                      fontFamily="monospace"
                      textAnchor="middle"
                      fill="#71717a"
                      className="font-bold uppercase tracking-wide"
                      transform={`rotate(-20, ${x}, ${svgHeight - paddingBottom + 14})`}
                    >
                      {d.label.length > 7 ? `${d.label.substring(0, 6)}.` : d.label}
                    </text>
                  </g>
                );
              })}

              {/* Area layout gradient background */}
              {chartType === "area" && spreadsheetChartData.length > 1 && (
                <path
                  d={spreadsheetChartData.map((d, idx) => {
                    const x = paddingLeft + stepX * idx;
                    const y = svgHeight - paddingBottom - ((d.val - minVal) / (maxVal - minVal)) * chartHeight;
                    return `${idx === 0 ? "M" : "L"} ${x} ${y}`;
                  }).join(" ") + ` L ${paddingLeft + stepX * (spreadsheetChartData.length - 1)} ${svgHeight - paddingBottom} L ${paddingLeft} ${svgHeight - paddingBottom} Z`}
                  fill="rgba(24, 24, 27, 0.06)"
                  stroke="none"
                />
              )}

              {/* Connected Lines */}
              {(chartType === "line" || chartType === "area") && spreadsheetChartData.length > 1 && (
                <path
                  d={spreadsheetChartData.map((d, idx) => {
                    const x = paddingLeft + stepX * idx;
                    const y = svgHeight - paddingBottom - ((d.val - minVal) / (maxVal - minVal)) * chartHeight;
                    return `${idx === 0 ? "M" : "L"} ${x} ${y}`;
                  }).join(" ")}
                  fill="none"
                  stroke="#18181b"
                  strokeWidth="3"
                  strokeLinecap="round"
                />
              )}
              
              {/* Baseline axis */}
              <line x1={paddingLeft} y1={svgHeight - paddingBottom} x2={svgWidth - paddingRight} y2={svgHeight - paddingBottom} stroke="#18181b" strokeWidth="2" />
            </svg>
          </div>
        )}
      </div>
    );
  };

  // Delete Fila (row) in grid
  const handleDeleteRow = (targetIdx: number) => {
    if (!activeSheet || activeSheet.grid.length <= 1) return;
    
    const updatedSheets = sheets.map((sh, sIdx) => {
      if (sIdx !== activeSheetIdx) return sh;
      
      const updatedGrid = sh.grid.filter((_, idx) => idx !== targetIdx);
      
      return {
        ...sh,
        grid: updatedGrid,
        rowCount: updatedGrid.length
      };
    });
    
    setSheets(updatedSheets);
  };

  // Dynamic Infinite Scroll: expands empty visual cells on the fly as the user scrolls near the bottom or right edge
  const handleGridScroll = (e: React.UIEvent<HTMLDivElement>) => {
    const { scrollTop, clientHeight, scrollHeight, scrollLeft, clientWidth, scrollWidth } = e.currentTarget;

    // Check if scrolled close to the bottom (within 90 pixels)
    if (scrollHeight - scrollTop - clientHeight < 90) {
      setVisibleRowsCount(prev => prev + 15);
    }

    // Check if scrolled close to the right edge (within 120 pixels)
    if (scrollWidth - scrollLeft - clientWidth < 120) {
      setVisibleColsCount(prev => prev + 4);
    }
  };

  // EXCEL COMPILATION WRITER: triggers XLSX library download directly to PC
  const handleDownloadSpreadsheet = () => {
    try {
      const wb = XLSX.utils.book_new();
      sheets.forEach(sh => {
        // Prune grid down to active non-empty bounding box for clean file outputs
        const cleanGrid = sh.grid.map(row => [...row]);
        const ws = XLSX.utils.aoa_to_sheet(cleanGrid);
        XLSX.utils.book_append_sheet(wb, ws, sh.sheetName);
      });
      
      const cleanName = doc.filename.replace(/\.[^/.]+$/, "");
      XLSX.writeFile(wb, `${cleanName || "hoja"}_modificado.xlsx`);
    } catch (err) {
      alert("Error exportando planilla: " + err);
    }
  };

  // Save selection position to restore it during color picker or dropdown selection interactions
  const saveSelection = () => {
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      const canvas = canvasRef.current || document.getElementById("rich-doc-canvas");
      if (canvas && canvas.contains(range.commonAncestorContainer)) {
        savedRangeRef.current = range.cloneRange();
        
        // Detect text indent of current selection
        let element = range.startContainer.nodeType === Node.ELEMENT_NODE 
          ? (range.startContainer as HTMLElement) 
          : range.startContainer.parentElement;
          
        while (element && element.id !== "rich-doc-canvas" && !["P", "DIV", "H1", "H2", "H3", "LI"].includes(element.tagName)) {
          element = element.parentElement;
        }
        
        if (element && element.id !== "rich-doc-canvas" && element.style.textIndent) {
          const match = element.style.textIndent.match(/^([\d.-]+)(cm|px|em)?$/);
          if (match) {
            setTextIndent(parseFloat(match[1]));
          }
        } else {
          setTextIndent(0);
        }
      }
    }
  };

  // Convert generic font-family class value (e.g. "font-serif") to beautiful font names
  const getFontFamilyString = (fontVal: string) => {
    if (fontVal.startsWith("uploaded-")) return `'${fontVal}', sans-serif`;
    switch (fontVal) {
      case "font-sans": return "'Inter', sans-serif";
      case "font-serif": return "'Merriweather', Georgia, serif";
      case "font-mono": return "'JetBrains Mono', monospace";
      case "font-display": return "'Space Grotesk', sans-serif";
      case "font-elegant": return "'Playfair Display', serif";
      case "font-classic": return "'Cinzel', serif";
      default: return fontVal;
    }
  };

  // Dynamic Page Number helper text based on format prefix settings and custom numbering criteria
  const getPageNumberText = (pageIdx: number = activePageIdx) => {
    // Hide if page numbers are disabled globally
    if (!showPageNumbers) return "";

    const currentPageNum = pageIdx + 1; // 1-based physical page

    // Hide if before start page index
    if (currentPageNum < pageNumberStartAtPageIdx) return "";

    // Hide if after end page index
    if (currentPageNum > pageNumberEndAtPageIdx) return "";

    // Hide on cover page (index 0) if exclude cover is set
    if (pageIdx === 0 && !pageNumberIncludeCover) return "";

    // Calculate the page number value to display
    let calculatedNum = pageNumberStartValue;
    if (pageNumberIncludeCover) {
      calculatedNum = pageIdx + pageNumberStartValue;
    } else {
      // If we exclude cover, physical page index 1 (the second page) is the first counted page
      calculatedNum = (pageIdx - 1) + pageNumberStartValue;
    }

    // Calculate total pages
    const totalPagesNum = pageNumberIncludeCover ? docPages.length : Math.max(1, docPages.length - 1);
    
    const prefix = pageNumberPrefix ? `${pageNumberPrefix} ` : "";
    
    switch (pageNumberFormat) {
      case "numero-solo":
        return `${calculatedNum}`;
      case "pagina-x":
        return `${prefix}${calculatedNum}`;
      case "pagina-x-de-y":
        return `${prefix}${calculatedNum} de ${totalPagesNum}`;
      case "x-y":
        return `${calculatedNum} / ${totalPagesNum}`;
      case "seccion-x":
        return `Sección I - ${prefix}${calculatedNum}`;
      default:
        return `${prefix}${calculatedNum} de ${totalPagesNum}`;
    }
  };

  // Exec Command format block helper for rich text Docx editing
  const handleRichCommand = (command: string, arg: string = "") => {
    // Restore selection first if any is stored
    if (savedRangeRef.current) {
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(savedRangeRef.current);
      }
    }
    document.execCommand(command, false, arg);
    const canvas = canvasRef.current || document.getElementById("rich-doc-canvas");
    if (canvas) {
      setEditedHtml(canvas.innerHTML);
    }
    saveSelection();
  };

  // Apply specific styles (such as colors, fonts, sizes) to the current text selection
  const applyStyleToSelection = (styleName: string, value: string) => {
    // Restore selection first if any is stored
    if (savedRangeRef.current) {
      const selection = window.getSelection();
      if (selection) {
        selection.removeAllRanges();
        selection.addRange(savedRangeRef.current);
      }
    }

    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const range = selection.getRangeAt(0);

    let finalValue = value;
    if (styleName === "font-family") {
      finalValue = getFontFamilyString(value);
    }
    
    if (range.collapsed) {
      // If cursor is sitting inside an element, set it directly on the parent style
      const parent = range.startContainer.parentElement;
      if (parent && parent.id !== "rich-doc-canvas") {
        parent.style.setProperty(styleName, finalValue);
      }
      return;
    }
    
    // Wrap selected nodes in a span styled with our property
    const span = document.createElement("span");
    span.style.setProperty(styleName, finalValue);
    try {
      range.surroundContents(span);
    } catch (e) {
      // Fallback formatting using standard browser commands with modern interceptions
      document.execCommand("styleWithCSS", false, "true");
      if (styleName === "color") {
        document.execCommand("foreColor", false, finalValue);
      } else if (styleName === "font-family") {
        document.execCommand("fontName", false, finalValue);
      } else if (styleName === "font-size") {
        const tempFontName = "temp-font-size-placeholder";
        document.execCommand("fontName", false, tempFontName);
        
        const canvas = canvasRef.current || document.getElementById("rich-doc-canvas");
        if (canvas) {
          const els = canvas.querySelectorAll(`font[face="${tempFontName}"], span[style*="${tempFontName}"]`);
          els.forEach((el: any) => {
            el.removeAttribute("face");
            el.style.fontFamily = "";
            el.style.fontSize = finalValue;
          });
        }
      }
    }
    
    const canvas = canvasRef.current || document.getElementById("rich-doc-canvas");
    if (canvas) {
      setEditedHtml(canvas.innerHTML);
    }
    
    saveSelection();
  };

  // Indentation commands: increases or decreases block left padding
  const handleIndentCommand = (direction: "increase" | "decrease") => {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const range = selection.getRangeAt(0);
    
    // Find closest block element
    let element = range.startContainer.nodeType === Node.ELEMENT_NODE 
      ? (range.startContainer as HTMLElement) 
      : range.startContainer.parentElement;
      
    while (element && element.id !== "rich-doc-canvas" && !["P", "DIV", "H1", "H2", "H3", "LI"].includes(element.tagName)) {
      element = element.parentElement;
    }
    
    if (element && element.id !== "rich-doc-canvas") {
      const currentPadding = parseInt(element.style.paddingLeft || "0", 10);
      const step = 20; // 20px padding increment
      const nextPadding = direction === "increase" 
        ? currentPadding + step 
        : Math.max(0, currentPadding - step);
      element.style.paddingLeft = nextPadding > 0 ? `${nextPadding}px` : "";
      
      const canvas = document.getElementById("rich-doc-canvas");
      if (canvas) {
        setEditedHtml(canvas.innerHTML);
      }
    } else {
      // Fallback standard command if no block element is found
      document.execCommand(direction === "increase" ? "indent" : "outdent", false);
      const canvas = document.getElementById("rich-doc-canvas");
      if (canvas) {
        setEditedHtml(canvas.innerHTML);
      }
    }
  };

  // Handle uploaded custom font files (.ttf, .otf, .woff, .woff2)
  const handleFontUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const result = event.target?.result as string;
      if (!result) return;

      // Create a unique css-safe font family name
      const cleanFontName = file.name
        .replace(/\.[^/.]+$/, "") // remove extension
        .replace(/[^a-zA-Z0-9]/g, "-") // replace non-alphanumeric with hyphens
        .toLowerCase();
      
      const fontNameWithPrefix = `uploaded-${cleanFontName}`;

      // Create style element to hold font-face
      const styleId = `font-style-${fontNameWithPrefix}`;
      let styleElement = document.getElementById(styleId) as HTMLStyleElement;
      if (!styleElement) {
        styleElement = document.createElement("style");
        styleElement.id = styleId;
        document.head.appendChild(styleElement);
      }

      // Determine format from file extension/type
      let format = "truetype";
      if (file.name.endsWith(".otf")) format = "opentype";
      else if (file.name.endsWith(".woff")) format = "woff";
      else if (file.name.endsWith(".woff2")) format = "woff2";

      styleElement.innerHTML = `
        @font-face {
          font-family: '${fontNameWithPrefix}';
          src: url('${result}') format('${format}');
          font-weight: normal;
          font-style: normal;
        }
      `;

      // Update uploadedFonts state
      setUploadedFonts(prev => {
        if (prev.some(f => f.cssName === fontNameWithPrefix)) return prev;
        return [...prev, { name: file.name.replace(/\.[^/.]+$/, ""), cssName: fontNameWithPrefix }];
      });

      // Set current font to the newly uploaded font
      setFontFamily(fontNameWithPrefix);
    };
    reader.readAsDataURL(file);
  };

  // Insert mathematical characters / arrows at cursor position inside rich contentEditable canvas
  const insertCharacterAtCaret = (char: string) => {
    const canvas = document.getElementById("rich-doc-canvas");
    if (!canvas) return;

    canvas.focus();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      canvas.innerHTML = canvas.innerHTML + char;
      setEditedHtml(canvas.innerHTML);
      return;
    }

    const range = selection.getRangeAt(0);
    let node = range.startContainer;
    let isInsideCanvas = false;
    while (node) {
      if (node === canvas) {
        isInsideCanvas = true;
        break;
      }
      node = node.parentNode as Node;
    }

    if (!isInsideCanvas) {
      canvas.innerHTML = canvas.innerHTML + char;
      setEditedHtml(canvas.innerHTML);
      return;
    }

    range.deleteContents();
    const textNode = document.createTextNode(char);
    range.insertNode(textNode);
    
    const newRange = document.createRange();
    newRange.setStartAfter(textNode);
    newRange.setEndAfter(textNode);
    selection.removeAllRanges();
    selection.addRange(newRange);

    setEditedHtml(canvas.innerHTML);
  };

  // Insert raw HTML at cursor position inside rich contentEditable canvas
  const insertHtmlAtCaret = (html: string) => {
    const canvas = canvasRef.current || document.getElementById("rich-doc-canvas");
    if (!canvas) return;

    canvas.focus();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      canvas.innerHTML = canvas.innerHTML + html;
      setEditedHtml(canvas.innerHTML);
      return;
    }

    const range = selection.getRangeAt(0);
    let node = range.startContainer;
    let isInsideCanvas = false;
    while (node) {
      if (node === canvas) {
        isInsideCanvas = true;
        break;
      }
      node = node.parentNode as Node;
    }

    if (!isInsideCanvas) {
      canvas.innerHTML = canvas.innerHTML + html;
      setEditedHtml(canvas.innerHTML);
      return;
    }

    range.deleteContents();
    
    const el = document.createElement("div");
    el.innerHTML = html;
    const fragment = document.createDocumentFragment();
    let lastNode: Node | null = null;
    while (el.firstChild) {
      lastNode = el.firstChild;
      fragment.appendChild(lastNode);
    }
    range.insertNode(fragment);
    
    if (lastNode) {
      const newRange = document.createRange();
      newRange.setStartAfter(lastNode);
      newRange.setEndAfter(lastNode);
      selection.removeAllRanges();
      selection.addRange(newRange);
    }

    setEditedHtml(canvas.innerHTML);
  };

  const getMathWrapper = (html: string, size: "chica" | "mediana" | "grande") => {
    const sizeStyle = 
      size === "chica" ? "font-size: 0.85em; margin: 0 2px;" :
      size === "grande" ? "font-size: 1.45em; margin: 0 6px;" :
      "font-size: 1.15em; margin: 0 4px;";
    return `<span style="display: inline-flex; align-items: center; vertical-align: middle; line-height: normal; ${sizeStyle}">${html}</span>`;
  };

  const getIntegralHtml = (definite: boolean, lower: string, upper: string, expression: string) => {
    if (!definite) {
      return `<span style="display: inline-flex; align-items: center; vertical-align: middle; font-family: 'Times New Roman', serif;"><span style="font-size: 1.3em; line-height: 1; font-style: normal; margin-right: 2px; vertical-align: middle;">&int;</span><span style="font-style: italic;">${expression}</span></span>`;
    }
    return `<span style="display: inline-flex; align-items: center; vertical-align: middle; font-family: 'Times New Roman', serif;"><span style="font-size: 1.4em; line-height: 1; position: relative; margin-right: 2px;">&int;</span><span style="display: inline-flex; flex-direction: column; font-size: 0.65em; line-height: 1; align-items: flex-start; margin-right: 4px;"><span style="margin-bottom: 1px; margin-top: -6px; font-weight: bold;">${upper}</span><span style="margin-top: 1px; font-weight: bold;">${lower}</span></span><span style="font-style: italic; margin-left: 1px;">${expression}</span></span>`;
  };

  const getRootHtml = (type: "square" | "cube" | "custom", index: string, expression: string) => {
    if (type === "square") {
      return `<span style="display: inline-flex; align-items: center; vertical-align: middle; font-family: 'Times New Roman', serif;"><span style="font-size: 1.2em; line-height: 1; margin-right: -1px; transform: scaleY(1.15);">&radic;</span><span style="border-top: 1.2px solid currentColor; padding-top: 1px; padding-left: 2px; font-style: italic;">${expression}</span></span>`;
    }
    const idx = type === "cube" ? "3" : index;
    return `<span style="display: inline-flex; align-items: center; vertical-align: middle; font-family: 'Times New Roman', serif;"><sup style="font-size: 0.65em; font-weight: bold; margin-right: -3px; margin-bottom: 7px; z-index: 1;">${idx}</sup><span style="font-size: 1.2em; line-height: 1; margin-right: -1px; transform: scaleY(1.15);">&radic;</span><span style="border-top: 1.2px solid currentColor; padding-top: 1px; padding-left: 2px; font-style: italic;">${expression}</span></span>`;
  };

  const getSumProdHtml = (type: "sum" | "prod", lower: string, upper: string, expression: string) => {
    const sym = type === "sum" ? "&Sigma;" : "&Pi;";
    return `<span style="display: inline-flex; align-items: center; vertical-align: middle; font-family: 'Times New Roman', serif;"><span style="display: inline-flex; flex-direction: column; align-items: center; font-size: 0.65em; line-height: 1; margin-right: 4px;"><span style="margin-bottom: 2px; font-weight: bold;">${upper}</span><span style="font-size: 1.85em; line-height: 1; margin: 1px 0; font-family: serif;">${sym}</span><span style="margin-top: 2px; font-weight: bold;">${lower}</span></span><span style="font-style: italic;">${expression}</span></span>`;
  };

  const getFractionHtml = (numerator: string, denominator: string) => {
    return `<span style="display: inline-flex; flex-direction: column; vertical-align: middle; text-align: center; font-family: 'Times New Roman', serif; line-height: 1; margin: 0 2px;"><span style="border-bottom: 1.2px solid currentColor; padding-bottom: 1.5px; font-style: italic; font-size: 0.9em; min-width: 12px; display: inline-block;">${numerator}</span><span style="padding-top: 1.5px; font-style: italic; font-size: 0.9em; min-width: 12px; display: inline-block;">${denominator}</span></span>`;
  };

  const getPowerHtml = (base: string, exponent: string) => {
    return `<span style="display: inline-flex; align-items: flex-start; vertical-align: middle; font-family: 'Times New Roman', serif;"><span style="font-style: italic;">${base}</span><sup style="font-size: 0.65em; font-weight: bold; margin-top: -0.25em; margin-left: 1px;">${exponent}</sup></span>`;
  };

  // Insert a fully editable beautiful cover page layout as a dedicated page 0
  const handleInsertCover = (type: "modern" | "editorial" | "tech" | "academic" | "custom") => {
    let coverHtml = "";
    const currentDate = new Date().toISOString().split('T')[0];

    switch (type) {
      case "custom":
        coverHtml = `
          <div class="portada-personalizada" style="padding: 3cm 2.5cm; text-align: center; width: 100% !important; height: 100% !important; min-height: 100% !important; box-sizing: border-box !important; margin: 0 !important; border-radius: 0 !important; display: flex; flex-direction: column; justify-content: space-between; ${customCoverStyleString || 'background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); border-left: 15px solid #2563eb; color: #1e3a8a;'}" contenteditable="true">
            <div>
              <p style="font-size: 11px; font-weight: 800; letter-spacing: 0.3em; text-transform: uppercase; margin-bottom: 24px;">LOREM IPSUM DOLOR</p>
              <h1 style="font-size: 38px; font-weight: 800; line-height: 1.15; letter-spacing: -0.03em; margin: 0 0 16px 0;">CONSECTETUR ADIPISCING</h1>
              <p style="font-size: 14px; max-width: 480px; font-weight: 400; line-height: 1.6; margin: 0 auto 40px auto;">Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident.</p>
            </div>
            <div style="border-top: 1px dashed currentColor; padding-top: 20px; display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; font-size: 11px; font-family: monospace;">
              <div>
                <strong style="display: block; margin-bottom: 4px; text-transform: uppercase;">Amet</strong>
                <span>Dolor</span>
              </div>
              <div>
                <strong style="display: block; margin-bottom: 4px; text-transform: uppercase;">Consectetur</strong>
                <span>Elit</span>
              </div>
              <div>
                <strong style="display: block; margin-bottom: 4px; text-transform: uppercase;">Tempus</strong>
                <span>${currentDate}</span>
              </div>
            </div>
          </div>
        `;
        break;

      case "modern":
        coverHtml = `
          <div class="portada-moderna" style="padding: 3cm 2.5cm; text-align: left; background: linear-gradient(135deg, #f4f4f5 0%, #e4e4e7 100%); width: 100% !important; height: 100% !important; min-height: 100% !important; box-sizing: border-box !important; margin: 0 !important; border-radius: 0 !important; font-family: 'Space Grotesk', 'Inter', sans-serif; border-left: 12px solid #059669; display: flex; flex-direction: column; justify-content: space-between;" contenteditable="true">
            <div>
              <p style="font-size: 11px; font-weight: 800; letter-spacing: 0.3em; color: #059669; text-transform: uppercase; margin-bottom: 24px;">LOREM IPSUM AMET</p>
              <h1 style="font-size: 38px; font-weight: 800; line-height: 1.15; letter-spacing: -0.03em; color: #18181b; margin: 0 0 16px 0;">CONSECTETUR ADIPISCING ELIT</h1>
              <p style="font-size: 14px; color: #4b5563; max-width: 480px; font-weight: 400; line-height: 1.6; margin-bottom: 40px;">Maecenas fringilla nisl id pellentesque vehicula. Mauris porttitor convallis sodales. Vestibulum ante ipsum primis in faucibus.</p>
            </div>
            <div style="border-top: 2px solid #d4d4d8; padding-top: 20px; display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; font-size: 11px; color: #52525b; font-family: monospace;">
              <div>
                <strong style="color: #18181b; display: block; margin-bottom: 4px; text-transform: uppercase;">Ametis</strong>
                <span>Lorem Ipsum</span>
              </div>
              <div>
                <strong style="color: #18181b; display: block; margin-bottom: 4px; text-transform: uppercase;">Doloris</strong>
                <span>Consectetur Elit</span>
              </div>
              <div>
                <strong style="color: #18181b; display: block; margin-bottom: 4px; text-transform: uppercase;">Tempus</strong>
                <span>${currentDate}</span>
              </div>
            </div>
          </div>
        `;
        break;

      case "editorial":
        coverHtml = `
          <div class="portada-editorial" style="padding: 3.5cm 2.5cm; text-align: center; background-color: #fdfbf7; border: 1px solid #e7e3d4; width: 100% !important; height: 100% !important; min-height: 100% !important; box-sizing: border-box !important; margin: 0 !important; border-radius: 0 !important; font-family: 'Playfair Display', 'Georgia', serif; display: flex; flex-direction: column; justify-content: space-between;" contenteditable="true">
            <div style="border-bottom: 1px solid #e7e3d4; padding-bottom: 20px;">
              <span style="font-size: 11px; font-weight: 600; letter-spacing: 0.25em; text-transform: uppercase; color: #857250;">LOREM IPSUM IV — AMET DOLOR</span>
            </div>
            <div style="margin: 40px 0;">
              <h1 style="font-size: 40px; font-weight: 400; line-height: 1.2; color: #2e281d; font-style: italic; margin-bottom: 15px;">CONSECTETUR ADIPISCING</h1>
              <div style="width: 40px; height: 1px; background-color: #857250; margin: 20px auto;"></div>
              <p style="font-family: 'Inter', sans-serif; font-size: 13px; color: #6b6455; max-width: 450px; margin: 0 auto; line-height: 1.7;">Aenean lacinia sollicitudin ante, id porttitor nunc. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum.</p>
            </div>
            <div style="font-family: 'Inter', sans-serif; font-size: 11px; color: #857250; letter-spacing: 0.1em; text-transform: uppercase; font-weight: 600;">
              <span>Nethelis Pressum</span>
              <span style="margin: 0 10px;">•</span>
              <span>Amet 2026</span>
            </div>
          </div>
        `;
        break;

      case "tech":
        coverHtml = `
          <div class="portada-tech" style="padding: 3cm 2.5cm; text-align: left; background-color: #0f172a; border: 1px solid #334155; width: 100% !important; height: 100% !important; min-height: 100% !important; box-sizing: border-box !important; margin: 0 !important; border-radius: 0 !important; font-family: 'JetBrains Mono', 'Fira Code', monospace; color: #94a3b8; display: flex; flex-direction: column; justify-content: space-between;" contenteditable="true">
            <div>
              <div style="display: inline-block; padding: 4px 10px; background-color: #1e293b; border: 1px solid #475569; border-radius: 6px; font-size: 9px; font-weight: bold; color: #38bdf8; margin-bottom: 24px;">
                LOREM: IPSUM_STABILE
              </div>
              <h1 style="font-size: 32px; font-weight: 800; line-height: 1.15; color: #f8fafc; letter-spacing: -0.02em; margin: 0 0 20px 0; font-family: sans-serif;">CONSECTETUR ADIPISCING V2</h1>
              <div style="width: 100%; height: 1px; background: linear-gradient(90deg, #38bdf8 0%, transparent 100%); margin-bottom: 25px;"></div>
              <p style="font-size: 12px; line-height: 1.6; max-width: 480px; color: #94a3b8; margin-bottom: 40px;">Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia curae; Aliquam erat volutpat.</p>
            </div>
            <div style="border-top: 1px solid #1e293b; padding-top: 20px; display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; font-size: 11px;">
              <div>
                <span style="color: #38bdf8; display: block; margin-bottom: 4px;">// LOREM_IPSUM</span>
                <span style="color: #f8fafc;">lorem_ipsum_core.json</span>
              </div>
              <div>
                <span style="color: #38bdf8; display: block; margin-bottom: 4px;">// AMET_CONSECTETUR</span>
                <span style="color: #f8fafc;">${currentDate} UTC</span>
              </div>
            </div>
          </div>
        `;
        break;

      case "academic":
        coverHtml = `
          <div class="portada-academica" style="padding: 3cm 2.5cm; text-align: center; background-color: #ffffff; border: 2px solid #1e3a8a; width: 100% !important; height: 100% !important; min-height: 100% !important; box-sizing: border-box !important; margin: 0 !important; border-radius: 0 !important; font-family: 'Times New Roman', Times, serif; display: flex; flex-direction: column; justify-content: space-between;" contenteditable="true">
            <div>
              <p style="font-size: 14px; font-weight: bold; color: #1e3a8a; text-transform: uppercase; margin-bottom: 4px; letter-spacing: 0.05em;">LOREM IPSUM ACADEMIA</p>
              <p style="font-size: 11px; color: #4b5563; margin-top: 0;">Consectetur Adipiscing Elit</p>
              <div style="width: 120px; height: 1.5px; background-color: #1e3a8a; margin: 15px auto 30px auto;"></div>
            </div>
            <div>
              <h1 style="font-size: 30px; font-weight: bold; color: #111827; line-height: 1.2; margin: 0 0 15px 0;">TEMPORA INCIDUNT UT LABORE</h1>
              <p style="font-size: 14px; font-style: italic; color: #4b5563; margin-bottom: 45px;">Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat</p>
            </div>
            <div style="max-width: 320px; margin: 0 auto; text-align: left; font-size: 12px; color: #374151; border-top: 1px solid #e5e7eb; padding-top: 20px; line-height: 1.8;">
              <div style="display: flex; justify-content: space-between;"><strong style="color: #111827;">Lorem:</strong> <span>Loremis Ipsumis</span></div>
              <div style="display: flex; justify-content: space-between;"><strong style="color: #111827;">Amet:</strong> <span>Ametis Doloris</span></div>
              <div style="display: flex; justify-content: space-between;"><strong style="color: #111827;">Ipsum:</strong> <span>K2021 — Elit Tempus</span></div>
              <div style="display: flex; justify-content: space-between;"><strong style="color: #111827;">Dolor:</strong> <span>${currentDate}</span></div>
            </div>
          </div>
        `;
        break;
    }

    setDocPages(prev => {
      // If the first page is completely empty, replace it, otherwise prepended as first page
      if (prev.length === 1 && (prev[0] === "" || prev[0] === "<p><br></p>")) {
        return [coverHtml];
      }
      return [coverHtml, ...prev];
    });
    setCoverHasMargins(false); // Make sure it takes up the full page (0cm margins)!
    setActivePageIdx(0);
    if (canvasRef.current) {
      canvasRef.current.innerHTML = coverHtml;
    }
    setEditedHtml(coverHtml);
  };

  // Insert structured cover page layout
  const handleInsertCoverPage = () => {
    const coverPageHtml = `
      <div class="portada-wrapper" style="border-bottom: 4px solid #18181b; padding: 3cm 2.5cm; width: 100% !important; height: 100% !important; min-height: 100% !important; box-sizing: border-box !important; margin: 0 !important; border-radius: 0 !important; text-align: left; background-color: #fafafa; font-family: 'Inter', sans-serif;" contenteditable="false">
        <p style="font-size: 11px; font-weight: bold; letter-spacing: 0.25em; color: #71717a; text-transform: uppercase; margin-bottom: 12px; font-family: monospace;">LOREM IPSUM ELOR_TT</p>
        <h1 style="font-size: 38px; font-weight: 800; line-height: 1.05; letter-spacing: -0.02em; color: #09090b; margin: 10px 0 20px 0;">CONSECTETUR ADIPISCING ELIT</h1>
        <div style="width: 70px; height: 3.5px; background-color: #18181b; margin-bottom: 25px;"></div>
        <p style="font-size: 14px; color: #4b5563; line-height: 1.5; max-width: 500px; margin-bottom: 40px;">Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident.</p>
        <div style="display: flex; gap: 40px; border-top: 1px solid #e4e4e7; padding-top: 24px; font-size: 11px; font-family: monospace; color: #71717a;">
          <div>
            <span style="font-weight: bold; color: #18181b; display: block; text-transform: uppercase; margin-bottom: 4px;">Lorem</span>
            <span>Loremis Ipsumis</span>
          </div>
          <div>
            <span style="font-weight: bold; color: #18181b; display: block; text-transform: uppercase; margin-bottom: 4px;">Ipsum</span>
            <span>2026-07-01</span>
          </div>
          <div>
            <span style="font-weight: bold; color: #18181b; display: block; text-transform: uppercase; margin-bottom: 4px;">Amet Elor_TT</span>
            <span>ELOR_TT-9842</span>
          </div>
        </div>
      </div>
      <hr style="border: none; border-bottom: 2px dashed #e4e4e7; margin: 40px 0; page-break-after: always;" />
    `;
    setEditedHtml(prev => coverPageHtml + prev);
    const canvas = document.getElementById("rich-doc-canvas");
    if (canvas) {
      canvas.innerHTML = coverPageHtml + canvas.innerHTML;
    }
  };

  // Insert abstract corporate business SVG chart
  const handleInsertAbstractChart = (chartStyle: "ventas" | "distribucion") => {
    let svgHtml = "";
    if (chartStyle === "ventas") {
      svgHtml = `
        <div contenteditable="false" style="text-align: center; margin: 24px auto; padding: 20px; background-color: #fafafa; border: 1px solid #e4e4e7; border-radius: 12px; max-width: 480px; font-family: 'Inter', sans-serif;" class="select-none">
          <h4 style="font-size: 13px; font-weight: bold; color: #18181b; margin-top: 0; margin-bottom: 12px;">Gráfico de Tendencia Elor_TT (Ventas Trimestrales)</h4>
          <svg viewBox="0 0 400 180" style="width: 100%; height: auto; max-height: 140px; display: block;">
            <line x1="30" y1="140" x2="380" y2="140" stroke="#e4e4e7" stroke-width="2" />
            <line x1="30" y1="20" x2="30" y2="140" stroke="#e4e4e7" stroke-width="2" />
            <path d="M30 110 L110 80 L190 100 L270 40 L350 50" fill="none" stroke="#18181b" stroke-width="3.5" stroke-linecap="round" />
            <path d="M30 110 L110 80 L190 100 L270 40 L350 50 L350 140 L30 140 Z" fill="rgba(24,24,27,0.05)" />
            <circle cx="30" cy="110" r="4.5" fill="#18181b" stroke="#ffffff" stroke-width="1.5" />
            <circle cx="110" cy="80" r="4.5" fill="#18181b" stroke="#ffffff" stroke-width="1.5" />
            <circle cx="190" cy="100" r="4.5" fill="#18181b" stroke="#ffffff" stroke-width="1.5" />
            <circle cx="270" cy="40" r="4.5" fill="#18181b" stroke="#ffffff" stroke-width="1.5" />
            <circle cx="350" cy="50" r="4.5" fill="#18181b" stroke="#ffffff" stroke-width="1.5" />
            <text x="30" y="156" font-size="8" font-family="monospace" text-anchor="middle" fill="#71717a">Trim 1</text>
            <text x="110" y="156" font-size="8" font-family="monospace" text-anchor="middle" fill="#71717a">Trim 2</text>
            <text x="190" y="156" font-size="8" font-family="monospace" text-anchor="middle" fill="#71717a">Trim 3</text>
            <text x="270" y="156" font-size="8" font-family="monospace" text-anchor="middle" fill="#71717a">Trim 4</text>
            <text x="350" y="156" font-size="8" font-family="monospace" text-anchor="middle" fill="#71717a">Trim 5</text>
          </svg>
          <p style="font-size: 9px; color: #a1a1aa; margin: 8px 0 0 0; font-family: monospace; letter-spacing: 0.1em; text-transform: uppercase;">Elor_TT Graphics Engine • Rendimiento General</p>
        </div>
      `;
    } else {
      svgHtml = `
        <div contenteditable="false" style="text-align: center; margin: 24px auto; padding: 20px; background-color: #fafafa; border: 1px solid #e4e4e7; border-radius: 12px; max-width: 480px; font-family: 'Inter', sans-serif;" class="select-none">
          <h4 style="font-size: 13px; font-weight: bold; color: #18181b; margin-top: 0; margin-bottom: 12px;">Gráfico de Distribución (Recursos Corporativos)</h4>
          <svg viewBox="0 0 400 180" style="width: 100%; height: auto; max-height: 140px; display: block;">
            <rect x="50" y="40" width="35" height="100" fill="#18181b" rx="4" />
            <rect x="130" y="70" width="35" height="70" fill="#3f3f46" rx="4" />
            <rect x="210" y="20" width="35" height="120" fill="#71717a" rx="4" />
            <rect x="290" y="90" width="35" height="50" fill="#a1a1aa" rx="4" />
            <line x1="30" y1="140" x2="370" y2="140" stroke="#e4e4e7" stroke-width="2" />
            <text x="67" y="156" font-size="8" font-family="monospace" text-anchor="middle" fill="#71717a">I+D</text>
            <text x="147" y="156" font-size="8" font-family="monospace" text-anchor="middle" fill="#71717a">Marketing</text>
            <text x="227" y="156" font-size="8" font-family="monospace" text-anchor="middle" fill="#71717a">Soporte</text>
            <text x="307" y="156" font-size="8" font-family="monospace" text-anchor="middle" fill="#71717a">Sistemas</text>
          </svg>
          <p style="font-size: 9px; color: #a1a1aa; margin: 8px 0 0 0; font-family: monospace; letter-spacing: 0.1em; text-transform: uppercase;">Elor_TT Graphics Engine • Distribución Relativa</p>
        </div>
      `;
    }

    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      const div = document.createElement("div");
      div.innerHTML = svgHtml;
      range.insertNode(div.firstElementChild!);
      const canvas = document.getElementById("rich-doc-canvas");
      if (canvas) {
        setEditedHtml(canvas.innerHTML);
      }
    } else {
      setEditedHtml(prev => prev + svgHtml);
    }
  };

  // Insert image via local file selection (fully reactive Base64 converter)
  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (event) => {
      const base64Data = event.target?.result as string;
      if (base64Data) {
        const imgHtml = `<img src="${base64Data}" alt="${file.name}" class="rounded-xl border border-zinc-200 shadow-sm max-w-full my-4 mx-auto" style="max-height: 280px; display: block;" />`;
        
        const selection = window.getSelection();
        if (selection && selection.rangeCount > 0) {
          const range = selection.getRangeAt(0);
          const div = document.createElement("div");
          div.innerHTML = imgHtml;
          range.insertNode(div.firstElementChild!);
          const canvas = document.getElementById("rich-doc-canvas");
          if (canvas) {
            setEditedHtml(canvas.innerHTML);
          }
        } else {
          setEditedHtml(prev => prev + imgHtml);
        }
      }
    };
    reader.readAsDataURL(file);
    e.target.value = ""; // Reset
  };

  const handleInsertImageUrl = () => {
    const url = prompt("Introduce la URL de la imagen:");
    if (!url) return;
    const imgHtml = `<img src="${url}" alt="Imagen insertada" class="rounded-xl border border-zinc-200 shadow-sm max-w-full my-4 mx-auto" style="max-height: 280px; display: block;" />`;
    
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      const div = document.createElement("div");
      div.innerHTML = imgHtml;
      range.insertNode(div.firstElementChild!);
      const canvas = document.getElementById("rich-doc-canvas");
      if (canvas) {
        setEditedHtml(canvas.innerHTML);
      }
    } else {
      setEditedHtml(prev => prev + imgHtml);
    }
  };

  // Download converted rich Word document (.doc) holding HTML formatting
  const handleDownloadWordDoc = () => {
    let pagesToUse = [...docPages];
    if (canvasRef.current) {
      pagesToUse[activePageIdx] = canvasRef.current.innerHTML;
    }
    const combinedHtml = pagesToUse.join("<hr style='page-break-after: always;' />");
    
    const header = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset="utf-8"><title>${doc.filename}</title><style>body { font-family: Calibri, Arial, sans-serif; }</style></head><body>`;
    const footer = "</body></html>";
    const compiledContent = header + combinedHtml + footer;
    
    const blob = new Blob([compiledContent], { type: "application/msword;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const cleanName = doc.filename.replace(/\.[^/.]+$/, "");
    
    a.href = url;
    a.download = `${cleanName || "documento"}_modificado.doc`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Download PDF content as plain text disguised as PDF file
  const handleDownloadPdf = () => {
    let text = "";
    if (doc.extension.toLowerCase() === ".pdf") {
      text = editedText;
    } else {
      let pagesToUse = [...docPages];
      if (canvasRef.current) {
        pagesToUse[activePageIdx] = canvasRef.current.innerHTML;
      }
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = pagesToUse.join("\n\n");
      text = tempDiv.innerText || tempDiv.textContent || "";
    }
    const blob = new Blob([text], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const cleanName = doc.filename.replace(/\.[^/.]+$/, "");
    
    a.href = url;
    a.download = `${cleanName || "documento"}_modificado.pdf`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Download Word/Text content as raw .txt file
  const handleDownloadPlainText = () => {
    let text = "";
    if (doc.extension.toLowerCase() === ".pdf") {
      text = editedText;
    } else {
      let pagesToUse = [...docPages];
      if (canvasRef.current) {
        pagesToUse[activePageIdx] = canvasRef.current.innerHTML;
      }
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = pagesToUse.join("\n\n--- PÁGINA ---\n\n");
      text = tempDiv.innerText || tempDiv.textContent || "";
    }

    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const cleanName = doc.filename.replace(/\.[^/.]+$/, "");
    
    a.href = url;
    a.download = `${cleanName || "documento"}_modificado.txt`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Highlight matches in PDF view
  const highlightedPdfExcerpt = useMemo(() => {
    if (!pdfSearch.trim()) {
      return (
        <textarea
          value={editedText}
          onChange={(e) => setEditedText(e.target.value)}
          className="w-full h-96 min-h-[350px] p-4 bg-gray-50 border border-gray-200 rounded-xl focus:outline-none focus:ring-1 focus:ring-black font-mono text-xs text-gray-800 leading-relaxed resize-y focus:bg-white transition-all"
          placeholder="Modifica el contenido del PDF aquí a tu gusto..."
        />
      );
    }

    const term = pdfSearch.trim().replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&"); // Escape regex chars
    const regex = new RegExp(`(${term})`, "gi");
    const parts = editedText.split(regex);

    return (
      <div className="flex flex-col gap-3">
        <div className="text-[10px] text-amber-800 font-bold uppercase tracking-wider bg-amber-50 border border-amber-200 p-2.5 rounded-lg select-none">
          ⚠️ Nota: Estás buscando términos en modo lectura. Para editar, elimina el término de búsqueda.
        </div>
        <div className="p-4 bg-gray-50 border border-gray-200 rounded-xl max-h-96 overflow-y-auto whitespace-pre-wrap text-xs font-mono text-gray-700 leading-relaxed leading-normal select-text">
          {parts.map((part, index) => 
            regex.test(part) ? (
              <mark key={index} className="bg-amber-100 text-amber-900 border-b border-amber-300 px-0.5 rounded-sm font-semibold">
                {part}
              </mark>
            ) : (
              part
            )
          )}
        </div>
      </div>
    );
  }, [editedText, pdfSearch]);

  const isExcel = [".xlsx", ".xls"].includes(doc.extension);
  const isPDF = doc.extension.toLowerCase() === ".pdf";
  const isWord = [".docx", ".doc"].includes(doc.extension);

  return (
    <div className="flex flex-col bg-white border border-gray-200 rounded-2xl overflow-hidden shadow-sm h-full" id="doc-preview-panel">
      
      {/* Panel main header info & action */}
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 border-b border-gray-200 px-6 py-4 bg-white">
        <div className="flex items-center gap-3">
          <div className="p-2.5 bg-black text-white rounded-lg select-none">
            {isExcel ? (
              <FileSpreadsheet className="w-5 h-5" />
            ) : (
              <FileText className="w-4.5 h-4.5" />
            )}
          </div>
          <div className="min-w-0">
            <h4 className="text-sm font-semibold text-gray-900 truncate max-w-[200px] sm:max-w-md">{doc.filename}</h4>
            <div className="flex items-center gap-2 mt-0.5 selection:bg-none">
              <span className="px-2 py-0.5 text-[10px] font-bold border rounded uppercase bg-gray-100 text-gray-800 border-gray-200">
                {doc.extension.substring(1)}
              </span>
              <span className="text-[10px] text-gray-400 font-bold uppercase tracking-widest font-mono">
                {formatFileSize(doc.size)}
              </span>
              <span className="text-[10px] text-zinc-700 font-bold uppercase tracking-widest bg-zinc-100 px-1.5 py-0.2 rounded border border-zinc-200">
                Elor_TT Engine
              </span>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2.5 w-full sm:w-auto">
          {/* Main Download Button */}
          {isExcel && (
            <button
              onClick={handleDownloadSpreadsheet}
              className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-4.5 py-2 px-4 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-full transition-all uppercase tracking-widest cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              Guardar en PC (.xlsx)
            </button>
          )}

          {isWord && (
            <div className="flex gap-2 w-full sm:w-auto">
              <button
                onClick={handleDownloadWordDoc}
                className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-full transition-all uppercase tracking-widest cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                Guardar en PC (.doc)
              </button>
              <button
                onClick={handleDownloadPlainText}
                className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-800 text-xs font-bold rounded-full transition-all uppercase tracking-widest border border-gray-200 cursor-pointer"
              >
                Txt Plano
              </button>
            </div>
          )}

          {isPDF && (
            <div className="flex gap-2 w-full sm:w-auto">
              <button
                onClick={handleDownloadPdf}
                className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-full transition-all uppercase tracking-widest cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                Guardar en PC (.pdf)
              </button>
              <button
                onClick={handleDownloadPlainText}
                className="flex-1 sm:flex-initial flex items-center justify-center gap-1.5 px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-800 text-xs font-bold rounded-full transition-all uppercase tracking-widest border border-gray-200 cursor-pointer"
              >
                Txt Plano
              </button>
            </div>
          )}

          <button
            onClick={onClear}
            className="flex items-center justify-center gap-1.5 p-2 bg-white hover:bg-gray-50 text-gray-500 hover:text-black rounded-lg transition-all border border-gray-200 cursor-pointer"
            title="Subir otro archivo"
          >
            <RefreshCw className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Main interactive preview and edit panel */}
      <div className="flex-1 overflow-auto p-5 select-text min-h-[450px]">
        {/* EXCEL GRID CONTROLLER */}
        {isExcel && activeSheet && (
          <div className="h-full flex flex-col gap-4">
            
            {/* Sheet Seleccion (Tabs) */}
            <div className="flex flex-wrap items-center gap-1.5 border-b border-gray-200 pb-2">
              {sheets.map((sheet, i) => (
                <button
                  key={sheet.sheetName}
                  onClick={() => {
                    setActiveSheetIdx(i);
                    setSpreadsheetSearch("");
                    setEditingCell(null);
                  }}
                  className={`px-3.5 py-2 text-[10px] font-bold uppercase tracking-wider rounded-lg transition-all border cursor-pointer
                    ${activeSheetIdx === i
                      ? "bg-black text-white border-black"
                      : "bg-white text-gray-500 border-gray-200 hover:bg-gray-50"
                    }`}
                >
                  <span className="flex items-center gap-1.5">
                    <FileSpreadsheet className="w-3.5 h-3.5" />
                    {sheet.sheetName}
                  </span>
                </button>
              ))}

              {/* [ + Nueva Hoja ] triggers form */}
              {isCreatingSheet ? (
                <form onSubmit={handleCreateSheet} className="flex items-center gap-1 bg-gray-50 border border-gray-200 p-1 rounded-lg">
                  <input
                    type="text"
                    placeholder="Nombre de pestaña..."
                    value={newSheetName}
                    onChange={(e) => setNewSheetName(e.target.value)}
                    autoFocus
                    className="px-2 py-1 text-xs bg-white border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-black min-w-[130px]"
                  />
                  <button
                    type="submit"
                    className="p-1 bg-black text-white rounded hover:bg-gray-800 transition"
                  >
                    <Check className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setIsCreatingSheet(false);
                      setNewSheetName("");
                    }}
                    className="p-1 bg-white border border-gray-300 text-gray-500 hover:text-black rounded transition"
                  >
                    ✕
                  </button>
                </form>
              ) : (
                <button
                  onClick={() => setIsCreatingSheet(true)}
                  className="px-3 py-2 text-[10px] font-bold uppercase tracking-wider text-black bg-white border border-dashed border-gray-300 hover:border-black rounded-lg transition-all cursor-pointer flex items-center gap-1"
                >
                  <Plus className="w-3.5 h-3.5" />
                  Nueva Hoja
                </button>
              )}
            </div>

            {/* Grid Editing Control Tools */}
            <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 bg-gray-50 border border-gray-200 p-3.5 rounded-2xl">
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-[10px] text-gray-400 font-bold uppercase tracking-widest font-mono mr-2">
                  Edición de Celdas:
                </span>
                <button
                  onClick={handleAddRow}
                  className="flex items-center gap-1 px-3 py-1.5 bg-white border border-gray-250 hover:border-black text-[10px] font-bold uppercase tracking-wider rounded-lg transition-all cursor-pointer"
                >
                  <Plus className="w-3 h-3 text-black" />
                  Agregar Fila
                </button>
                <button
                  onClick={handleAddCol}
                  className="flex items-center gap-1 px-3 py-1.5 bg-white border border-gray-250 hover:border-black text-[10px] font-bold uppercase tracking-wider rounded-lg transition-all cursor-pointer"
                >
                  <Plus className="w-3 h-3 text-black" />
                  Agregar Columna
                </button>

                <div className="w-px h-5 bg-gray-300 mx-1.5 hidden sm:block" />

                {/* Grow Infinite Visor count quickly */}
                <button
                  onClick={() => setVisibleRowsCount(prev => prev + 100)}
                  className="px-2.5 py-1.5 bg-white border border-gray-200 hover:border-black text-[10px] font-bold uppercase tracking-wider rounded-lg transition-all cursor-pointer"
                  title="Expande los límites visibles para seguir completando"
                >
                  +100 Filas Visor
                </button>
                <button
                  onClick={() => setVisibleColsCount(prev => prev + 15)}
                  className="px-2.5 py-1.5 bg-white border border-gray-200 hover:border-black text-[10px] font-bold uppercase tracking-wider rounded-lg transition-all cursor-pointer"
                  title="Expande los límites de columnas visibles"
                >
                  +15 Cols Visor
                </button>

                <div className="text-[9.5px] text-gray-500 font-mono pl-2 border-l border-gray-200">
                  Guardados: <span className="text-black font-semibold">{activeSheet.grid.length}F</span> x <span className="text-black font-semibold">{activeSheet.columnCount}C</span>
                </div>
              </div>

              {/* Advanced Real-Time Search coordinates with next/prev highlight navigation */}
              <div className="flex items-center gap-2">
                {/* Excel Search Box */}
                <div className="relative w-full md:w-56">
                  <Search className="absolute left-2.5 top-2.5 w-3.5 h-3.5 text-gray-400" />
                  <input
                    type="text"
                    placeholder="Buscar en esta hoja..."
                    value={spreadsheetSearch}
                    onChange={(e) => {
                      setSpreadsheetSearch(e.target.value);
                      setEditingCell(null);
                    }}
                    className="w-full bg-white border border-gray-200 text-xs rounded-xl pl-8 pr-3 py-2 focus:outline-none focus:ring-1 focus:ring-black hover:border-gray-300 transition-all font-sans"
                  />
                </div>

                {/* Matched counter layout */}
                {spreadsheetSearch.trim() !== "" && (
                  <div className="flex items-center gap-1 px-3 py-1 bg-yellow-55 bg-amber-50 border border-yellow-200 text-yellow-905 rounded-xl text-xs font-mono select-none whitespace-nowrap">
                    {searchMatches.length > 0 ? (
                      <div className="flex items-center gap-2">
                        <span className="font-bold text-gray-800">{currentMatchIdx + 1}/{searchMatches.length}</span>
                        <div className="flex gap-0.5">
                          <button
                            type="button"
                            onClick={() => setCurrentMatchIdx(prev => (prev - 1 + searchMatches.length) % searchMatches.length)}
                            className="p-0.5 hover:bg-gray-200 rounded text-black font-bold transition cursor-pointer"
                            title="Anterior"
                          >
                            <ChevronUp className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => setCurrentMatchIdx(prev => (prev + 1) % searchMatches.length)}
                            className="p-0.5 hover:bg-gray-200 rounded text-black font-bold transition cursor-pointer"
                            title="Siguiente"
                          >
                            <ChevronDown className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                    ) : (
                      <span className="text-gray-400 text-[10px] font-bold uppercase">Sin coincidencias</span>
                    )}
                  </div>
                )}
              </div>
            </div>

            {/* Hint message & dynamic info */}
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 text-[10px] text-gray-400 uppercase tracking-widest font-mono select-none">
              <span className="flex items-center gap-1">
                <Grid className="w-3.5 h-3.5 text-black" />
                <span>Modo Grilla Infinita: Doble click sobre cualquier celda para completar un valor</span>
              </span>
              <span className="text-zinc-600 font-bold bg-zinc-105 text-zinc-750 px-2.5 py-0.5 rounded-full border border-zinc-200">
                Nethels Industries
              </span>
            </div>

            {/* Structured Cell Table Grid */}
            <div 
              onScroll={handleGridScroll}
              className="flex-1 overflow-auto border border-gray-200 rounded-2xl max-h-[550px] bg-white shadow-inner relative"
            >
              <table className="w-full border-collapse text-left text-xs text-gray-700 table-fixed">
                <thead>
                  <tr className="bg-gray-50/80 backdrop-blur-sm sticky top-0 border-b border-gray-200 divide-x divide-gray-200 select-none z-10">
                    <th className="w-12 text-center text-[9px] uppercase tracking-widest font-mono py-2.5 px-1 font-bold text-gray-400">
                      #
                    </th>
                    {Array.from({ length: visibleColsCount }).map((_, colIdx) => (
                      <th key={colIdx} className="text-center text-[10px] uppercase tracking-wider font-mono py-2 px-3 text-gray-500 w-[120px] font-bold">
                        {getColLetter(colIdx)}
                      </th>
                    ))}
                    <th className="w-10 text-center text-[9px] uppercase tracking-widest font-mono py-2 px-1 text-gray-400">
                      Acc.
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100 divide-x divide-gray-100 bg-white leading-normal">
                  {Array.from({ length: visibleRowsCount }).map((_, rIdx) => {
                    // Fetch cell record or fall back empty
                    const rowData = activeSheet.grid[rIdx];
                    
                    return (
                      <tr key={rIdx} className="hover:bg-gray-50/50 transition-colors">
                        {/* Excel index row number */}
                        <td className="bg-gray-50/70 text-[10px] text-center font-mono py-2 px-1 font-semibold text-gray-400 border-r border-gray-200 sticky left-0 z-1 select-none">
                          {rIdx + 1}
                        </td>
                        
                        {/* Columns values */}
                        {Array.from({ length: visibleColsCount }).map((_, colIdx) => {
                          const cellVal = rowData ? rowData[colIdx] : "";
                          const isEditing = editingCell?.rowIdx === rIdx && editingCell?.colIdx === colIdx;
                          
                          // Determine real-time search match status inside rendered cells
                          const term = spreadsheetSearch.toLowerCase().trim();
                          const isMatch = term !== "" && String(cellVal ?? "").toLowerCase().includes(term);
                          
                          const activeMatchObj = searchMatches[currentMatchIdx];
                          const isActiveMatch = isMatch && activeMatchObj && activeMatchObj.rowIdx === rIdx && activeMatchObj.colIdx === colIdx;
                          
                          return (
                            <td 
                              key={colIdx} 
                              id={`cell-${rIdx}-${colIdx}`}
                              onDoubleClick={() => handleStartEditing(rIdx, colIdx, cellVal)}
                              className={`py-1.5 px-2 whitespace-nowrap overflow-hidden text-ellipsis w-[120px] border-r border-gray-100 text-gray-800 relative group cursor-pointer hover:bg-yellow-50/40 transition-shadow duration-150
                                ${isActiveMatch 
                                  ? "bg-amber-300 ring-2 ring-black font-extrabold text-black z-10" 
                                  : isMatch 
                                    ? "bg-yellow-101 bg-yellow-100 text-yellow-900 border-b border-yellow-300 font-semibold" 
                                    : ""
                                }`}
                              title="Haz doble clic para editar"
                            >
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editValue}
                                  onChange={(e) => setEditValue(e.target.value)}
                                  onBlur={() => handleFinishEditing(rIdx, colIdx)}
                                  onKeyDown={(e) => {
                                    if (e.key === "Enter") handleFinishEditing(rIdx, colIdx);
                                    if (e.key === "Escape") setEditingCell(null);
                                  }}
                                  autoFocus
                                  className="absolute inset-0 w-full h-full bg-white px-2 border-2 border-black focus:outline-none font-sans text-xs text-black shadow-lg z-20"
                                />
                              ) : cellVal !== undefined && cellVal !== null && cellVal !== "" ? (
                                <span className={typeof cellVal === "number" ? "font-mono font-medium text-black" : ""}>
                                  {String(cellVal)}
                                </span>
                              ) : (
                                <span className="text-gray-200 font-mono italic select-none text-[10.5px]">-</span>
                              )}
                            </td>
                          );
                        })}

                        {/* Actions column */}
                        <td className="text-center py-1.5 px-1 bg-gray-50/30">
                          <button
                            onClick={() => handleDeleteRow(rIdx)}
                            disabled={!rowData || activeSheet.grid.length <= 1}
                            className="p-1 hover:text-rose-600 text-gray-400 transition-colors disabled:opacity-30 disabled:hover:text-gray-400 cursor-pointer"
                            title="Eliminar Fila"
                          >
                            <Trash2 className="w-3.5 h-3.5 mx-auto" strokeWidth={2.5} />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Word Document (.docx / .doc) or PDF Interactive High-Fidelity Editor */}
        {(isWord || isPDF) && (
          <div className="h-full flex flex-col gap-4 text-left">
            
            {/* TOGGLABLE HIGH-FIDELITY TOOLBAR (Simple / Expanded states) */}
            {!isToolbarExpanded ? (
              <div className="bg-[#161925]/60 backdrop-blur-md border border-zinc-800 p-2 rounded-2xl flex items-center justify-between select-none shadow-md transition-all duration-300">
                <div className="flex items-center gap-2.5 pl-3">
                  <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                  <span className="text-[11px] font-mono font-bold text-zinc-400 uppercase tracking-widest">Editor de Documentos Listo</span>
                </div>
                <button
                  type="button"
                  onClick={() => setIsToolbarExpanded(true)}
                  className="px-4 py-1.5 bg-[#1E2230] hover:bg-[#2E3344] text-white rounded-full text-[10px] font-bold font-mono uppercase tracking-widest flex items-center gap-2 transition-all duration-250 border border-zinc-750 cursor-pointer shadow active:scale-95"
                >
                  <Sparkles className="w-3.5 h-3.5 text-amber-400" />
                  Abrir Opciones de Diseño
                </button>
              </div>
            ) : (
              <div className="bg-zinc-50 border border-zinc-250 p-3.5 rounded-2xl flex flex-col gap-3 select-none shadow-md transition-all duration-300">
                
                {/* Toolbar Control Header with Back Button */}
                <div className="flex items-center justify-between border-b border-zinc-200 pb-2.5">
                  <div className="flex items-center gap-2">
                    <div className="w-2.5 h-2.5 rounded-full bg-black animate-ping shrink-0" />
                    <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-widest font-mono">Panel de Configuración de Documento</span>
                  </div>
                  <button
                    type="button"
                    onClick={() => setIsToolbarExpanded(false)}
                    className="px-3.5 py-1.5 bg-zinc-200 hover:bg-zinc-300 text-zinc-800 hover:text-black rounded-lg text-[10px] font-bold font-mono uppercase tracking-wider transition-all duration-200 cursor-pointer flex items-center gap-1.5 active:scale-95 shadow-sm border border-zinc-300/40"
                  >
                    <span>← Volver Atrás</span>
                  </button>
                </div>

                {/* Horizontal Tabs / Triggers */}
                <div className="flex items-center gap-1.5 border-b border-zinc-200 pb-2">
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => toggleSection("inicio")}
                    className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${openSections.inicio ? "bg-zinc-900 text-white" : "text-zinc-600 hover:bg-zinc-200/60"}`}
                  >
                    Inicio
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => toggleSection("insertar")}
                    className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${openSections.insertar ? "bg-zinc-900 text-white" : "text-zinc-600 hover:bg-zinc-200/60"}`}
                  >
                    Insertar
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => toggleSection("pagina")}
                    className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${openSections.pagina ? "bg-zinc-900 text-white" : "text-zinc-600 hover:bg-zinc-200/60"}`}
                  >
                    Página
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => toggleSection("simbolos")}
                    className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${openSections.simbolos ? "bg-zinc-900 text-white" : "text-zinc-600 hover:bg-zinc-200/60"}`}
                  >
                    Símbolos
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => toggleSection("configPag")}
                    className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${openSections.configPag ? "bg-zinc-900 text-white" : "text-zinc-600 hover:bg-zinc-200/60"}`}
                  >
                    Config Pág.
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => toggleSection("ajustes")}
                    className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${openSections.ajustes ? "bg-zinc-900 text-white" : "text-zinc-600 hover:bg-zinc-200/60"}`}
                  >
                    Ajustes
                  </button>
                </div>

                {/* Active Tab Content Area - Horizontal, Compact, Rows with Columns */}
                <div className="animate-fade-in text-zinc-800">
                  
                  {/* SECTION 1: INICIO */}
                  {openSections.inicio && (
                    <div className="grid grid-cols-1 lg:grid-cols-12 gap-3.5 text-xs py-1 items-stretch">
                      
                      {/* Column 1: Format tools & Alignment (Left, Center, Right) */}
                      <div className="lg:col-span-4 flex flex-wrap items-center gap-1.5 bg-zinc-100/70 border border-zinc-200/80 rounded-xl p-2 shadow-xs">
                        {/* Negrita, Itálica, Subrayado */}
                        <div className="flex items-center gap-0.5 bg-white border border-zinc-200 rounded-lg p-0.5 shadow-xs shrink-0">
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleRichCommand("bold")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Negrita"
                          >
                            <Bold className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleRichCommand("italic")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Itálica"
                          >
                            <Italic className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleRichCommand("underline")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Subrayado"
                          >
                            <Underline className="w-3.5 h-3.5" />
                          </button>
                        </div>

                        {/* Headings */}
                        <div className="flex items-center gap-0.5 bg-white border border-zinc-200 rounded-lg p-0.5 shadow-xs shrink-0">
                          {(["h1", "h2", "h3"] as const).map(h => (
                            <button
                              key={h}
                              type="button"
                              onMouseDown={(e) => e.preventDefault()}
                              onClick={() => handleRichCommand("formatBlock", h)}
                              className="px-2 py-1 hover:bg-zinc-100 rounded text-zinc-800 font-bold text-[10px] uppercase transition active:scale-95 cursor-pointer"
                              title={h.toUpperCase()}
                            >
                              {h.toUpperCase()}
                            </button>
                          ))}
                        </div>

                        {/* Alignment buttons group (Alinear izquierda, centro, derecha) */}
                        <div className="flex bg-white border border-zinc-200 rounded-lg p-0.5 shadow-xs shrink-0" title="Alineación">
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleRichCommand("justifyLeft")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Alinear a la Izquierda"
                          >
                            <AlignLeft className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleRichCommand("justifyCenter")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Centrar"
                          >
                            <AlignCenter className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleRichCommand("justifyRight")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Alinear a la Derecha"
                          >
                            <AlignRight className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>

                      {/* Column 2: TIPOGRAFÍA Y COLOR (TODO BIEN CERCA) */}
                      <div className="lg:col-span-5 flex flex-wrap items-center gap-2.5 bg-emerald-50/50 border border-emerald-150 rounded-xl p-2 shadow-xs">
                        {/* Ámbito de Letra (Scope) */}
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] font-bold text-emerald-850 uppercase tracking-wider">Aplicar Letra a:</span>
                          <select
                            value={typographyScope}
                            onChange={(e) => setTypographyScope(e.target.value as "documento" | "seleccion")}
                            className="bg-white border border-emerald-300 text-[10px] rounded-lg px-2 py-1 font-bold text-emerald-950 focus:outline-none shadow-2xs"
                          >
                            <option value="seleccion">✨ Selección</option>
                            <option value="documento">📄 Todo el Archivo</option>
                          </select>
                        </div>

                        {/* Selector de Fuente */}
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] font-bold text-emerald-800 uppercase tracking-wider">Tipo de Letra:</span>
                          <select
                            value={fontFamily}
                            onChange={(e) => handleFontFamilyChange(e.target.value)}
                            className="bg-white border border-emerald-200 text-[11px] rounded-lg px-2 py-1 font-semibold text-zinc-800 focus:outline-none shadow-xs"
                          >
                            <option value="font-sans">Inter</option>
                            <option value="font-serif">Merriweather</option>
                            <option value="font-mono">JetBrains Mono</option>
                            <option value="font-display">Space Grotesk</option>
                            <option value="font-elegant">Playfair Display</option>
                            <option value="font-classic">Cinzel</option>
                            {uploadedFonts.map(font => (
                              <option key={font.cssName} value={font.cssName}>
                                {font.name}
                              </option>
                            ))}
                          </select>
                        </div>

                        {/* Tamaño de Letra */}
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] font-bold text-emerald-800 uppercase tracking-wider">Tamaño:</span>
                          <select
                            value={fontSize}
                            onChange={(e) => handleFontSizeChange(e.target.value)}
                            className="bg-white border border-emerald-200 text-[11px] rounded-lg px-2 py-1 font-semibold text-zinc-800 focus:outline-none shadow-xs w-16"
                          >
                            <option value="12px">12px</option>
                            <option value="13px">13px</option>
                            <option value="15px">15px</option>
                            <option value="17px">17px</option>
                            <option value="20px">20px</option>
                            <option value="24px">24px</option>
                            <option value="28px">28px</option>
                            <option value="32px">32px</option>
                          </select>
                        </div>

                        {/* Importar Letra (.ttf) */}
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] font-bold text-emerald-800 uppercase tracking-wider">Importar Letra:</span>
                          <div className="relative">
                            <label
                              htmlFor="font-file-upload-accordion"
                              onMouseDown={(e) => e.preventDefault()}
                              className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[10px] font-bold flex items-center gap-1 transition cursor-pointer select-none shadow-sm border border-emerald-700/30"
                            >
                              <Upload className="w-3 h-3 text-emerald-100" />
                              Subir .ttf
                            </label>
                            <input
                              type="file"
                              id="font-file-upload-accordion"
                              accept=".ttf,.otf,.woff,.woff2"
                              onChange={handleFontUpload}
                              className="hidden"
                            />
                          </div>
                        </div>

                        {/* Color de Letra */}
                        <div className="flex flex-col gap-0.5" title="Color de Letra">
                          <span className="text-[9px] font-bold text-emerald-800 uppercase tracking-wider">Color Letra:</span>
                          <div className="flex items-center gap-1.5 bg-white border border-emerald-200 rounded-lg px-1.5 py-0.5 shadow-xs h-[26px]">
                            {/* Color display indicator */}
                            <div 
                              className="w-4 h-4 rounded-full border border-zinc-200 shadow-inner"
                              style={{ backgroundColor: textColor }}
                            />
                            {/* Native picker with beautiful choose button trigger */}
                            <div className="relative">
                              <button
                                type="button"
                                className="px-1.5 py-0.5 bg-zinc-100 hover:bg-zinc-200 border border-zinc-300 rounded text-[9px] font-bold text-zinc-700 transition"
                              >
                                Elegir
                              </button>
                              <input 
                                type="color" 
                                value={textColor}
                                onChange={(e) => handleTextColorChange(e.target.value)} 
                                className="absolute inset-0 opacity-0 cursor-pointer w-full h-full scale-110"
                              />
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Column 3: Copiar/Pegar Estilo, Listas & Sangrías */}
                      <div className="lg:col-span-3 flex flex-wrap items-center gap-1.5 bg-zinc-100/70 border border-zinc-200/80 rounded-xl p-2 shadow-xs">
                        {/* Lists & Indents */}
                        <div className="flex items-center gap-0.5 bg-white border border-zinc-200 rounded-lg p-0.5 shadow-xs shrink-0">
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleRichCommand("insertUnorderedList")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Viñetas"
                          >
                            <List className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleRichCommand("insertOrderedList")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Lista Numerada"
                          >
                            <ListOrdered className="w-3.5 h-3.5" />
                          </button>
                          <div className="w-px h-4 bg-zinc-200 mx-1"></div>
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleIndentCommand("decrease")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Reducir Sangría"
                          >
                            <Outdent className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => handleIndentCommand("increase")}
                            className="p-1.5 hover:bg-zinc-100 rounded text-zinc-800 transition active:scale-95 cursor-pointer"
                            title="Aumentar Sangría"
                          >
                            <Indent className="w-3.5 h-3.5" />
                          </button>
                        </div>

                        {/* Copiar/Pegar Estilos */}
                        <div className="flex items-center gap-1">
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={handleCopyFormat}
                            className="px-2 py-1.5 bg-white hover:bg-zinc-100 border border-zinc-200 rounded-lg text-[9px] font-bold text-zinc-700 flex items-center gap-1 transition active:scale-95 cursor-pointer shadow-xs"
                            title="Copiar Estilo"
                          >
                            <Copy className="w-3 h-3 text-zinc-500" />
                            Copiar
                          </button>
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={handlePasteFormat}
                            disabled={!copiedFormat}
                            className={`px-2 py-1.5 border rounded-lg text-[9px] font-bold flex items-center gap-1 transition active:scale-95 cursor-pointer shadow-xs
                              ${copiedFormat 
                                ? "bg-zinc-900 text-white hover:bg-zinc-800 border-zinc-900" 
                                : "bg-zinc-50 text-zinc-400 border-zinc-200 cursor-not-allowed"}`}
                            title="Pegar Estilo"
                          >
                            <Paintbrush className="w-3 h-3" />
                            Pegar
                          </button>
                        </div>
                      </div>

                    </div>
                  )}

                  {/* SECTION 2: INSERTAR */}
                  {openSections.insertar && (
                    <div className="flex items-center gap-2 flex-wrap text-xs py-1">
                      <button
                        type="button"
                        onClick={() => setShowCoverSelectorModal(true)}
                        className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white font-bold rounded-xl text-[10px] flex items-center gap-1.5 transition cursor-pointer shadow-xs uppercase tracking-wider"
                      >
                        <Sparkles className="w-3.5 h-3.5 text-amber-300 animate-pulse" />
                        Configurar e Insertar Portada...
                      </button>

                      <button
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => handleInsertAbstractChart("ventas")}
                        className="px-3 py-1.5 bg-white hover:bg-zinc-100 border border-zinc-200 rounded-lg text-[10px] font-bold text-zinc-800 flex items-center gap-1 transition cursor-pointer active:scale-95 shadow-xs"
                      >
                        Gráfico Ventas
                      </button>
                      
                      <button
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => handleInsertAbstractChart("distribucion")}
                        className="px-3 py-1.5 bg-white hover:bg-zinc-100 border border-zinc-200 rounded-lg text-[10px] font-bold text-zinc-800 flex items-center gap-1 transition cursor-pointer active:scale-95 shadow-xs"
                      >
                        Gráfico Distribución
                      </button>

                      <div className="relative">
                        <label
                          htmlFor="local-img-upload-accordion"
                          onMouseDown={(e) => e.preventDefault()}
                          className="px-3 py-1.5 bg-white hover:bg-zinc-100 border border-zinc-200 rounded-lg text-[10px] font-bold text-zinc-800 flex items-center gap-1 transition cursor-pointer select-none shadow-xs"
                        >
                          Imagen Local
                        </label>
                        <input
                          type="file"
                          id="local-img-upload-accordion"
                          accept="image/*"
                          onChange={handleImageUpload}
                          className="hidden"
                        />
                      </div>

                      <button
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={handleInsertImageUrl}
                        className="px-3 py-1.5 bg-white hover:bg-zinc-100 border border-zinc-200 rounded-lg text-[10px] font-bold text-zinc-800 flex items-center gap-1 transition cursor-pointer active:scale-95 shadow-xs"
                      >
                        Imagen URL
                      </button>
                    </div>
                  )}

                  {/* SECTION 3: PÁGINA */}
                  {openSections.pagina && (
                    <div className="grid grid-cols-1 md:grid-cols-12 gap-3 text-xs py-1 items-stretch animate-fade-in text-zinc-800">
                      
                      {/* Sub-column 1: Diseño de Hoja */}
                      <div className="md:col-span-3 flex flex-col gap-1.5 bg-amber-50/40 border border-amber-200/65 rounded-xl p-2.5 shadow-xs">
                        <span className="text-[9px] font-bold text-amber-800 uppercase tracking-wider font-mono">Diseño de Hoja</span>
                        
                        {/* Selector de Tipo de Página */}
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[10px] text-zinc-600 font-medium">Tamaño de Hoja:</span>
                          <select
                            value={pageSize}
                            onChange={(e) => {
                              const size = e.target.value as any;
                              setPageSize(size);
                              if (size === "A4") { setCustomWidth(21.0); setCustomHeight(29.7); }
                              else if (size === "A5") { setCustomWidth(14.8); setCustomHeight(21.0); }
                              else if (size === "Carta") { setCustomWidth(21.6); setCustomHeight(27.9); }
                            }}
                            className="bg-white border border-zinc-250 text-[11px] rounded px-1.5 py-1 font-semibold text-zinc-800 focus:outline-none"
                          >
                            <option value="A4">A4 (21 x 29.7 cm)</option>
                            <option value="A5">A5 (14.8 x 21 cm)</option>
                            <option value="Carta">Carta (21.6 x 27.9 cm)</option>
                            <option value="Infinita">♾️ Página Infinita (Ancho Fijo)</option>
                            <option value="Personalizado">⚙️ Personalizado (cm)</option>
                          </select>
                        </div>

                        {/* Custom size inputs */}
                        {(pageSize === "Personalizado" || pageSize === "Infinita") && (
                          <div className="grid grid-cols-2 gap-1.5 mt-0.5 animate-fade-in">
                            <div className="flex flex-col gap-0.5">
                              <span className="text-[9px] text-zinc-500">Ancho (cm):</span>
                              <input
                                type="number"
                                min={10}
                                max={50}
                                step={0.1}
                                value={customWidth}
                                onChange={(e) => setCustomWidth(parseFloat(e.target.value) || 21)}
                                className="bg-white border border-zinc-250 rounded px-1.5 py-0.5 font-semibold text-zinc-850"
                              />
                            </div>
                            {pageSize === "Personalizado" && (
                              <div className="flex flex-col gap-0.5 animate-fade-in">
                                <span className="text-[9px] text-zinc-500">Alto (cm):</span>
                                <input
                                  type="number"
                                  min={10}
                                  max={120}
                                  step={0.1}
                                  value={customHeight}
                                  onChange={(e) => setCustomHeight(parseFloat(e.target.value) || 29.7)}
                                  className="bg-white border border-zinc-250 rounded px-1.5 py-0.5 font-semibold text-zinc-850"
                                />
                              </div>
                            )}
                          </div>
                        )}

                        {/* Color de Hoja */}
                        <div className="flex items-center justify-between gap-1.5 mt-1">
                          <span className="text-[10px] font-medium text-zinc-600">Fondo de Hoja:</span>
                          <div className="flex items-center gap-1 bg-white border border-zinc-250 rounded-lg px-2 py-0.5 shadow-xs h-[22px]">
                            <div className="relative flex items-center justify-center w-3 h-3 rounded-full border border-zinc-300 overflow-hidden" style={{ backgroundColor: paperBg }}>
                              <input 
                                type="color" 
                                value={paperBg}
                                onChange={(e) => setPaperBg(e.target.value)} 
                                className="absolute inset-0 opacity-0 cursor-pointer w-full h-full scale-150"
                              />
                            </div>
                            <span className="text-[9px] font-mono font-bold text-zinc-500 uppercase">{paperBg}</span>
                          </div>
                        </div>

                        {/* Color de Letra Base */}
                        <div className="flex items-center justify-between gap-1.5">
                          <span className="text-[10px] font-medium text-zinc-600">Color Base:</span>
                          <div className="flex items-center gap-1 bg-white border border-zinc-250 rounded-lg px-2 py-0.5 shadow-xs h-[22px]">
                            <div className="relative flex items-center justify-center w-3 h-3 rounded-full border border-zinc-300 overflow-hidden" style={{ backgroundColor: globalTextColor }}>
                              <input 
                                type="color" 
                                value={globalTextColor}
                                onChange={(e) => setGlobalTextColor(e.target.value)} 
                                className="absolute inset-0 opacity-0 cursor-pointer w-full h-full scale-150"
                              />
                            </div>
                            <span className="text-[9px] font-mono font-bold text-zinc-500 uppercase">{globalTextColor}</span>
                          </div>
                        </div>
                      </div>

                      {/* Sub-column 2: Márgenes de Página */}
                      <div className="md:col-span-3 flex flex-col gap-1.5 bg-blue-50/30 border border-blue-150 rounded-xl p-2.5 shadow-xs">
                        <span className="text-[9px] font-bold text-blue-800 uppercase tracking-wider font-mono flex items-center justify-between">
                          <span>Márgenes (cm)</span>
                          <span className="text-[8px] bg-blue-100 px-1 rounded text-blue-700">Reglas Activas</span>
                        </span>

                        {/* Margins Mode */}
                        <div className="flex bg-white border border-zinc-200 rounded p-0.5 shrink-0 mt-0.5">
                          <button
                            type="button"
                            onClick={() => {
                              setMarginsMode("2m");
                              setMarginRight(marginLeft);
                              setMarginBottom(marginTop);
                            }}
                            className={`flex-1 text-[9px] font-bold py-0.5 rounded transition ${marginsMode === "2m" ? "bg-blue-600 text-white" : "text-zinc-600 hover:bg-zinc-150"}`}
                            title="Muestra 2 márgenes: arriba e izquierda perpendiculares"
                          >
                            2 Marg. L/T
                          </button>
                          <button
                            type="button"
                            onClick={() => setMarginsMode("4m")}
                            className={`flex-1 text-[9px] font-bold py-0.5 rounded transition ${marginsMode === "4m" ? "bg-blue-600 text-white" : "text-zinc-600 hover:bg-zinc-150"}`}
                            title="Personaliza los 4 márgenes de forma independiente"
                          >
                            4 Márgenes
                          </button>
                        </div>

                        {/* Margins inputs */}
                        <div className="grid grid-cols-2 gap-1.5 mt-0.5">
                          <div className="flex flex-col gap-0.5">
                            <span className="text-[9px] text-zinc-500">Superior (cm):</span>
                            <input
                              type="number"
                              min={0.5}
                              max={10}
                              step={0.1}
                              value={marginTop}
                              onChange={(e) => {
                                const val = parseFloat(e.target.value) || 2.5;
                                setMarginTop(val);
                                if (marginsMode === "2m") setMarginBottom(val);
                              }}
                              className="bg-white border border-zinc-250 rounded px-1 py-0.5 font-mono text-[10px] font-bold text-zinc-800"
                            />
                          </div>
                          <div className="flex flex-col gap-0.5">
                            <span className="text-[9px] text-zinc-500">Izquierdo (cm):</span>
                            <input
                              type="number"
                              min={0.5}
                              max={10}
                              step={0.1}
                              value={marginLeft}
                              onChange={(e) => {
                                const val = parseFloat(e.target.value) || 2.5;
                                setMarginLeft(val);
                                if (marginsMode === "2m") setMarginRight(val);
                              }}
                              className="bg-white border border-zinc-250 rounded px-1 py-0.5 font-mono text-[10px] font-bold text-zinc-800"
                            />
                          </div>

                          {marginsMode === "4m" && (
                            <>
                              <div className="flex flex-col gap-0.5 animate-fade-in">
                                <span className="text-[9px] text-zinc-500">Derecho (cm):</span>
                                <input
                                  type="number"
                                  min={0.5}
                                  max={10}
                                  step={0.1}
                                  value={marginRight}
                                  onChange={(e) => setMarginRight(parseFloat(e.target.value) || 2.5)}
                                  className="bg-white border border-zinc-250 rounded px-1 py-0.5 font-mono text-[10px] font-bold text-zinc-800"
                                />
                              </div>
                              {pageDimensions.height && (
                                <div className="flex flex-col gap-0.5 animate-fade-in">
                                  <span className="text-[9px] text-zinc-500">Inferior (cm):</span>
                                  <input
                                    type="number"
                                    min={0.5}
                                    max={10}
                                    step={0.1}
                                    value={marginBottom}
                                    onChange={(e) => setMarginBottom(parseFloat(e.target.value) || 2.5)}
                                    className="bg-white border border-zinc-250 rounded px-1 py-0.5 font-mono text-[10px] font-bold text-zinc-800"
                                  />
                                </div>
                              )}
                            </>
                          )}
                        </div>

                        {/* Ver Líneas Toggle */}
                        <label className="flex items-center gap-1.5 font-mono text-[9px] text-zinc-600 font-bold cursor-pointer select-none mt-1">
                          <input
                            type="checkbox"
                            checked={showRulerLines}
                            onChange={(e) => setShowRulerLines(e.target.checked)}
                            className="accent-blue-600 rounded h-3 w-3"
                          />
                          Mostrar Guías de Regla
                        </label>
                      </div>

                      {/* Sub-column 3: Encabezado y Pie de Página */}
                      <div className="md:col-span-3 flex flex-col gap-1.5 bg-zinc-150/40 border border-zinc-200/80 rounded-xl p-2.5 shadow-xs">
                        <span className="text-[9px] font-bold text-zinc-600 uppercase tracking-wider font-mono">Encabezado y Pie</span>
                        
                        {/* Encabezado row */}
                        <div className="flex flex-col gap-1">
                          <label className="flex items-center gap-1 text-[10px] font-bold text-zinc-700 cursor-pointer select-none">
                            <input
                              type="checkbox"
                              checked={showHeader}
                              onChange={(e) => setShowHeader(e.target.checked)}
                              className="accent-black rounded h-3 w-3"
                            />
                            <span>Encabezado</span>
                          </label>
                          {showHeader && (
                            <div className="flex items-center gap-1 animate-fade-in">
                              <input
                                type="text"
                                value={docHeader}
                                onChange={(e) => setDocHeader(e.target.value)}
                                placeholder="Texto..."
                                className="flex-1 bg-white border border-zinc-250 text-[10px] rounded px-1.5 py-0.5 focus:outline-none text-zinc-850 font-medium"
                              />
                              <div className="relative w-3.5 h-3.5 rounded-full border border-zinc-350 overflow-hidden shrink-0" style={{ backgroundColor: headerColor }}>
                                <input
                                  type="color"
                                  value={headerColor}
                                  onChange={(e) => setHeaderColor(e.target.value)}
                                  className="absolute inset-0 opacity-0 cursor-pointer w-full h-full scale-150"
                                />
                              </div>
                            </div>
                          )}
                        </div>

                        {/* Pie de Página row */}
                        <div className="flex flex-col gap-1 mt-1">
                          <label className="flex items-center gap-1 text-[10px] font-bold text-zinc-700 cursor-pointer select-none">
                            <input
                              type="checkbox"
                              checked={showFooter}
                              onChange={(e) => setShowFooter(e.target.checked)}
                              className="accent-black rounded h-3 w-3"
                            />
                            <span>Pie de Página</span>
                          </label>
                          {showFooter && (
                            <div className="flex items-center gap-1 animate-fade-in">
                              <input
                                type="text"
                                value={docFooter}
                                onChange={(e) => setDocFooter(e.target.value)}
                                placeholder="Texto..."
                                className="flex-1 bg-white border border-zinc-250 text-[10px] rounded px-1.5 py-0.5 focus:outline-none text-zinc-850 font-medium"
                              />
                              <div className="relative w-3.5 h-3.5 rounded-full border border-zinc-350 overflow-hidden shrink-0" style={{ backgroundColor: footerColor }}>
                                <input
                                  type="color"
                                  value={footerColor}
                                  onChange={(e) => setFooterColor(e.target.value)}
                                  className="absolute inset-0 opacity-0 cursor-pointer w-full h-full scale-150"
                                />
                              </div>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* Sub-column 4: Numeración Ultra Personalizable */}
                      <div className="md:col-span-3 flex flex-col gap-1.5 bg-emerald-50/50 border border-emerald-150 rounded-xl p-2.5 shadow-xs">
                        <span className="text-[9px] font-bold text-emerald-800 uppercase tracking-wider font-mono">Numeración</span>
                        
                        <div className="flex items-center gap-2">
                          <label className="flex items-center gap-1.5 font-mono text-[10px] text-emerald-900 font-bold cursor-pointer select-none">
                            <input
                              type="checkbox"
                              checked={showPageNumbers}
                              onChange={(e) => setShowPageNumbers(e.target.checked)}
                              className="accent-emerald-700 rounded h-3 w-3"
                            />
                            Activar Numeración
                          </label>
                        </div>

                        {showPageNumbers && (
                          <div className="flex flex-col gap-1.5 animate-fade-in text-[10px]">
                            <div className="flex items-center justify-between gap-1">
                              <span className="text-zinc-500">Formato:</span>
                              <select
                                value={pageNumberFormat}
                                onChange={(e) => setPageNumberFormat(e.target.value)}
                                className="bg-white border border-emerald-200 rounded px-1 py-0.5 focus:outline-none font-semibold text-zinc-800 text-[9px] w-[95px]"
                              >
                                <option value="numero-solo">Número Solo</option>
                                <option value="pagina-x">Pág. X</option>
                                <option value="pagina-x-de-y">Pág. X de Y</option>
                                <option value="x-y">X / Y</option>
                                <option value="seccion-x">Sección I - Pág. X</option>
                              </select>
                            </div>

                            <div className="flex items-center justify-between gap-1">
                              <span className="text-zinc-500">Prefijo:</span>
                              <input
                                type="text"
                                value={pageNumberPrefix}
                                onChange={(e) => setPageNumberPrefix(e.target.value)}
                                placeholder="Ej: Pág."
                                className="bg-white border border-emerald-200 rounded px-1.5 py-0.5 focus:outline-none font-semibold text-zinc-850 text-[9px] w-[95px]"
                              />
                            </div>

                            <div className="flex items-center justify-between gap-1">
                              <span className="text-zinc-500">Alineación:</span>
                              <div className="flex bg-white border border-emerald-200 rounded p-0.5 shrink-0">
                                {[
                                  { value: "left", label: "Izq" },
                                  { value: "center", label: "Cent" },
                                  { value: "right", label: "Der" }
                                ].map(align => (
                                  <button
                                    key={align.value}
                                    type="button"
                                    onClick={() => setPageNumberAlign(align.value as any)}
                                    className={`px-1.5 py-0.5 text-[8.5px] font-bold rounded transition capitalize cursor-pointer
                                      ${pageNumberAlign === align.value 
                                        ? "bg-emerald-700 text-white shadow-xs" 
                                        : "text-zinc-600 hover:bg-emerald-50"}`}
                                  >
                                    {align.label}
                                  </button>
                                ))}
                              </div>
                            </div>

                            <div className="border-t border-emerald-100/70 my-1 pt-1.5 flex flex-col gap-1.5">
                              <label className="flex items-center gap-1.5 text-[9.5px] text-zinc-600 font-bold cursor-pointer select-none">
                                <input
                                  type="checkbox"
                                  checked={pageNumberIncludeCover}
                                  onChange={(e) => setPageNumberIncludeCover(e.target.checked)}
                                  className="accent-emerald-700 rounded h-3 w-3"
                                />
                                Incluir portada en conteo
                              </label>

                              <div className="flex items-center justify-between gap-1 mt-0.5">
                                <span className="text-zinc-500" title="Número inicial que se muestra">Número Inicial:</span>
                                <input
                                  type="number"
                                  min={1}
                                  value={pageNumberStartValue}
                                  onChange={(e) => setPageNumberStartValue(Math.max(1, parseInt(e.target.value) || 1))}
                                  className="bg-white border border-emerald-200 rounded px-1.5 py-0.5 focus:outline-none font-semibold text-zinc-850 text-[9px] w-[50px] text-center"
                                />
                              </div>

                              <div className="flex items-center justify-between gap-1">
                                <span className="text-zinc-500" title="Fila física en la que empieza la numeración">Ver desde pág:</span>
                                <input
                                  type="number"
                                  min={1}
                                  value={pageNumberStartAtPageIdx}
                                  onChange={(e) => setPageNumberStartAtPageIdx(Math.max(1, parseInt(e.target.value) || 1))}
                                  className="bg-white border border-emerald-200 rounded px-1.5 py-0.5 focus:outline-none font-semibold text-zinc-850 text-[9px] w-[50px] text-center"
                                />
                              </div>

                              <div className="flex items-center justify-between gap-1">
                                <span className="text-zinc-500" title="Límite máximo de numeración">Hasta pág:</span>
                                <input
                                  type="number"
                                  min={1}
                                  value={pageNumberEndAtPageIdx}
                                  onChange={(e) => setPageNumberEndAtPageIdx(Math.max(1, parseInt(e.target.value) || 999))}
                                  className="bg-white border border-emerald-200 rounded px-1.5 py-0.5 focus:outline-none font-semibold text-zinc-850 text-[9px] w-[50px] text-center"
                                />
                              </div>
                            </div>
                          </div>
                        )}
                      </div>

                    </div>
                  )}

                  {/* SECTION 4: SÍMBOLOS */}
                  {openSections.simbolos && (
                    <div className="flex flex-col gap-3.5 text-xs py-1 animate-fade-in text-zinc-800">
                      
                      {/* SUB-SECCIÓN 1: SÍMBOLOS RÁPIDOS */}
                      <div className="bg-zinc-50 border border-zinc-200 rounded-xl p-2.5 shadow-xs flex flex-col gap-2">
                        <span className="text-[9px] font-bold text-zinc-500 uppercase tracking-wider font-mono">Inserción Rápida de Símbolos y Números</span>
                        
                        <div className="grid grid-cols-1 lg:grid-cols-4 gap-2">
                          {/* Álgebra y Letras Griegas */}
                          <div className="bg-white border border-zinc-200/80 rounded-lg p-2 flex flex-col gap-1.5">
                            <span className="text-[8.5px] font-extrabold text-emerald-800 font-mono uppercase">Álgebra & Griego</span>
                            <div className="flex flex-wrap gap-1">
                              {["π", "θ", "±", "≠", "≈", "∞", "×", "÷", "≤", "≥", "α", "β", "γ", "λ", "Δ", "μ", "σ", "δ", "ε", "φ", "ω"].map(sym => (
                                <button
                                  key={sym}
                                  type="button"
                                  onMouseDown={(e) => e.preventDefault()}
                                  onClick={() => insertCharacterAtCaret(sym)}
                                  className="w-6 h-6 flex items-center justify-center hover:bg-zinc-900 hover:text-white border border-zinc-200 rounded text-[11px] font-bold font-mono transition bg-zinc-50 text-zinc-700 cursor-pointer active:scale-90"
                                >
                                  {sym}
                                </button>
                              ))}
                            </div>
                          </div>

                          {/* Flechas y Operadores */}
                          <div className="bg-white border border-zinc-200/80 rounded-lg p-2 flex flex-col gap-1.5">
                            <span className="text-[8.5px] font-extrabold text-blue-800 font-mono uppercase">Flechas & Relaciones</span>
                            <div className="flex flex-wrap gap-1">
                              {["→", "←", "↑", "↓", "↔", "⇒", "⇐", "⇔", "⇄", "➔", "∝", "≡", "≅", "∼", "∴", "∵", "✓", "✗"].map(sym => (
                                <button
                                  key={sym}
                                  type="button"
                                  onMouseDown={(e) => e.preventDefault()}
                                  onClick={() => insertCharacterAtCaret(sym)}
                                  className="w-6 h-6 flex items-center justify-center hover:bg-zinc-900 hover:text-white border border-zinc-200 rounded text-[11px] font-bold font-mono transition bg-zinc-50 text-zinc-700 cursor-pointer active:scale-90"
                                >
                                  {sym}
                                </button>
                              ))}
                            </div>
                          </div>

                          {/* Lógica y Conjuntos */}
                          <div className="bg-white border border-zinc-200/80 rounded-lg p-2 flex flex-col gap-1.5">
                            <span className="text-[8.5px] font-extrabold text-amber-800 font-mono uppercase">Lógica & Conjuntos</span>
                            <div className="flex flex-wrap gap-1">
                              {["∀", "∃", "∈", "∉", "⊂", "⊃", "⊆", "⊇", "∪", "∩", "∅", "¬", "∧", "∨", "⊕", "⊗", "°", "ℯ"].map(sym => (
                                <button
                                  key={sym}
                                  type="button"
                                  onMouseDown={(e) => e.preventDefault()}
                                  onClick={() => insertCharacterAtCaret(sym)}
                                  className="w-6 h-6 flex items-center justify-center hover:bg-zinc-900 hover:text-white border border-zinc-200 rounded text-[11px] font-bold font-mono transition bg-zinc-50 text-zinc-700 cursor-pointer active:scale-90"
                                >
                                  {sym}
                                </button>
                              ))}
                            </div>
                          </div>

                          {/* Números & Formato Exponentes */}
                          <div className="bg-white border border-zinc-200/80 rounded-lg p-2 flex flex-col gap-1.5">
                            <span className="text-[8.5px] font-extrabold text-purple-800 font-mono uppercase">Sub / Super Índices Rápidos</span>
                            <div className="flex flex-col gap-1.5 justify-between h-full pb-0.5">
                              {/* Quick subscript superscripts */}
                              <div className="flex flex-wrap gap-1">
                                {[
                                  { label: "x²", html: "<sup>2</sup>" },
                                  { label: "x³", html: "<sup>3</sup>" },
                                  { label: "xⁿ", html: "<sup>n</sup>" },
                                  { label: "x₁", html: "<sub>1</sub>" },
                                  { label: "x₂", html: "<sub>2</sub>" },
                                  { label: "xᵢ", html: "<sub>i</sub>" }
                                ].map((item, idx) => (
                                  <button
                                    key={idx}
                                    type="button"
                                    onMouseDown={(e) => e.preventDefault()}
                                    onClick={() => insertHtmlAtCaret(item.html)}
                                    className="px-1.5 h-6 flex items-center justify-center hover:bg-zinc-900 hover:text-white border border-zinc-200 rounded text-[10px] font-bold font-mono transition bg-zinc-50 text-zinc-700 cursor-pointer active:scale-90"
                                    title={`Insertar ${item.label}`}
                                  >
                                    {item.label}
                                  </button>
                                ))}
                              </div>
                              {/* Native selection script wrapper */}
                              <div className="flex gap-1.5 mt-1">
                                <button
                                  type="button"
                                  onMouseDown={(e) => e.preventDefault()}
                                  onClick={() => handleRichCommand("superscript")}
                                  className="flex-1 py-1 px-1.5 bg-zinc-100 hover:bg-zinc-200 border border-zinc-300 rounded text-[9px] font-bold text-zinc-700 transition active:scale-95"
                                  title="Convertir selección en Superíndice (exponente)"
                                >
                                  x<sup>y</sup> Exponente
                                </button>
                                <button
                                  type="button"
                                  onMouseDown={(e) => e.preventDefault()}
                                  onClick={() => handleRichCommand("subscript")}
                                  className="flex-1 py-1 px-1.5 bg-zinc-100 hover:bg-zinc-200 border border-zinc-300 rounded text-[9px] font-bold text-zinc-700 transition active:scale-95"
                                  title="Convertir selección en Subíndice"
                                >
                                  x<sub>y</sub> Subíndice
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* SUB-SECCIÓN 2: CONSTRUCTOR INTERACTIVO DE FÓRMULAS */}
                      <div className="bg-emerald-50/40 border border-emerald-200/80 rounded-xl p-3 shadow-xs flex flex-col gap-2">
                        <div className="flex justify-between items-center border-b border-emerald-150 pb-1.5">
                          <span className="text-[10px] font-bold text-emerald-900 uppercase tracking-wider font-mono">Constructor de Fórmulas Matemáticas Avanzadas</span>
                          <span className="text-[9px] font-medium text-emerald-700 bg-emerald-100 px-2 py-0.5 rounded-full font-mono">Personalizable & WYSIWYG</span>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-12 gap-3.5 items-stretch">
                          {/* Col 1: Tipo de Fórmula */}
                          <div className="md:col-span-3 flex flex-col gap-1.5 bg-white border border-emerald-150 rounded-lg p-2">
                            <span className="text-[8px] font-extrabold text-emerald-800 uppercase tracking-wide font-mono">1. Estructura base</span>
                            
                            <div className="flex flex-col gap-1">
                              <button
                                type="button"
                                onMouseDown={(e) => e.preventDefault()}
                                onClick={() => setMathType("integral")}
                                className={`w-full text-left px-2.5 py-1.5 rounded-md font-semibold text-[11px] flex items-center justify-between transition
                                  ${mathType === "integral" 
                                    ? "bg-emerald-700 text-white shadow-xs" 
                                    : "bg-zinc-50 text-zinc-700 hover:bg-zinc-100 border border-zinc-200/70"}`}
                              >
                                <span>Integral</span>
                                <span className="font-mono text-xs">∫ dx</span>
                              </button>

                              <button
                                type="button"
                                onMouseDown={(e) => e.preventDefault()}
                                onClick={() => setMathType("raiz")}
                                className={`w-full text-left px-2.5 py-1.5 rounded-md font-semibold text-[11px] flex items-center justify-between transition
                                  ${mathType === "raiz" 
                                    ? "bg-emerald-700 text-white shadow-xs" 
                                    : "bg-zinc-50 text-zinc-700 hover:bg-zinc-100 border border-zinc-200/70"}`}
                              >
                                <span>Radical / Raíz</span>
                                <span className="font-mono text-xs">√x</span>
                              </button>

                              <button
                                type="button"
                                onMouseDown={(e) => e.preventDefault()}
                                onClick={() => setMathType("sumatoria")}
                                className={`w-full text-left px-2.5 py-1.5 rounded-md font-semibold text-[11px] flex items-center justify-between transition
                                  ${mathType === "sumatoria" 
                                    ? "bg-emerald-700 text-white shadow-xs" 
                                    : "bg-zinc-50 text-zinc-700 hover:bg-zinc-100 border border-zinc-200/70"}`}
                              >
                                <span>Sumatoria / Prod.</span>
                                <span className="font-mono text-xs">Σ / Π</span>
                              </button>

                              <button
                                type="button"
                                onMouseDown={(e) => e.preventDefault()}
                                onClick={() => setMathType("fraccion")}
                                className={`w-full text-left px-2.5 py-1.5 rounded-md font-semibold text-[11px] flex items-center justify-between transition
                                  ${mathType === "fraccion" 
                                    ? "bg-emerald-700 text-white shadow-xs" 
                                    : "bg-zinc-50 text-zinc-700 hover:bg-zinc-100 border border-zinc-200/70"}`}
                              >
                                <span>Fracción</span>
                                <span className="font-mono text-xs">x/y</span>
                              </button>

                              <button
                                type="button"
                                onMouseDown={(e) => e.preventDefault()}
                                onClick={() => setMathType("potencia")}
                                className={`w-full text-left px-2.5 py-1.5 rounded-md font-semibold text-[11px] flex items-center justify-between transition
                                  ${mathType === "potencia" 
                                    ? "bg-emerald-700 text-white shadow-xs" 
                                    : "bg-zinc-50 text-zinc-700 hover:bg-zinc-100 border border-zinc-200/70"}`}
                              >
                                <span>Potencia / Exponente</span>
                                <span className="font-mono text-xs">xⁿ</span>
                              </button>
                            </div>
                          </div>

                          {/* Col 2: Parámetros del constructor */}
                          <div className="md:col-span-5 flex flex-col gap-2.5 bg-white border border-emerald-150 rounded-lg p-2.5 justify-between">
                            <div>
                              <span className="text-[8px] font-extrabold text-emerald-800 uppercase tracking-wide font-mono block mb-2">2. Configurar variables y extremos</span>

                              {/* INTEGRAL CONTROLS */}
                              {mathType === "integral" && (
                                <div className="flex flex-col gap-2 animate-fade-in text-[11px]">
                                  <div className="flex items-center gap-4">
                                    <span className="text-zinc-500 font-bold">Tipo:</span>
                                    <label className="flex items-center gap-1 cursor-pointer font-semibold text-zinc-700">
                                      <input 
                                        type="radio" 
                                        checked={integralDefinite} 
                                        onChange={() => setIntegralDefinite(true)} 
                                        className="accent-emerald-700 h-3.5 w-3.5"
                                      />
                                      Definida (Con límites)
                                    </label>
                                    <label className="flex items-center gap-1 cursor-pointer font-semibold text-zinc-700">
                                      <input 
                                        type="radio" 
                                        checked={!integralDefinite} 
                                        onChange={() => setIntegralDefinite(false)} 
                                        className="accent-emerald-700 h-3.5 w-3.5"
                                      />
                                      Indefinida
                                    </label>
                                  </div>

                                  {integralDefinite && (
                                    <div className="grid grid-cols-2 gap-2 mt-0.5">
                                      <div className="flex flex-col gap-0.5">
                                        <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Límite Inferior:</span>
                                        <input 
                                          type="text" 
                                          value={integralLower} 
                                          onChange={(e) => setIntegralLower(e.target.value)} 
                                          onFocus={() => setActiveFormulaInput("integralLower")}
                                          placeholder="a, 0, x..."
                                          className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                        />
                                      </div>
                                      <div className="flex flex-col gap-0.5">
                                        <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Límite Superior:</span>
                                        <input 
                                          type="text" 
                                          value={integralUpper} 
                                          onChange={(e) => setIntegralUpper(e.target.value)} 
                                          onFocus={() => setActiveFormulaInput("integralUpper")}
                                          placeholder="b, ∞, n..."
                                          className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                        />
                                      </div>
                                    </div>
                                  )}

                                  <div className="flex flex-col gap-0.5">
                                    <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Función a Integrar:</span>
                                    <input 
                                      type="text" 
                                      value={integralExpression} 
                                      onChange={(e) => setIntegralExpression(e.target.value)} 
                                      onFocus={() => setActiveFormulaInput("integralExpression")}
                                      placeholder="f(x) dx, x² dx..."
                                      className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                    />
                                  </div>
                                </div>
                              )}

                              {/* ROOT / RADICAL CONTROLS */}
                              {mathType === "raiz" && (
                                <div className="flex flex-col gap-2 animate-fade-in text-[11px]">
                                  <div className="flex items-center gap-3">
                                    <span className="text-zinc-500 font-bold">Índice:</span>
                                    <label className="flex items-center gap-1 cursor-pointer font-semibold text-zinc-700">
                                      <input 
                                        type="radio" 
                                        checked={rootType === "square"} 
                                        onChange={() => setRootType("square")} 
                                        className="accent-emerald-700 h-3.5 w-3.5"
                                      />
                                      Cuadrada (√)
                                    </label>
                                    <label className="flex items-center gap-1 cursor-pointer font-semibold text-zinc-700">
                                      <input 
                                        type="radio" 
                                        checked={rootType === "cube"} 
                                        onChange={() => setRootType("cube")} 
                                        className="accent-emerald-700 h-3.5 w-3.5"
                                      />
                                      Cúbica (³√)
                                    </label>
                                    <label className="flex items-center gap-1 cursor-pointer font-semibold text-zinc-700">
                                      <input 
                                        type="radio" 
                                        checked={rootType === "custom"} 
                                        onChange={() => setRootType("custom")} 
                                        className="accent-emerald-700 h-3.5 w-3.5"
                                      />
                                      Personalizada (ⁿ√)
                                    </label>
                                  </div>

                                  {rootType === "custom" && (
                                    <div className="flex flex-col gap-0.5 animate-fade-in mt-0.5">
                                      <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider font-mono">Índice de la raíz (número o letra):</span>
                                      <input 
                                        type="text" 
                                        value={rootIndex} 
                                        onChange={(e) => setRootIndex(e.target.value)} 
                                        onFocus={() => setActiveFormulaInput("rootIndex")}
                                        placeholder="n, 4, 5, k..."
                                        className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                      />
                                    </div>
                                  )}

                                  <div className="flex flex-col gap-0.5">
                                    <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Expresión (Radicando):</span>
                                    <input 
                                      type="text" 
                                      value={rootExpression} 
                                      onChange={(e) => setRootExpression(e.target.value)} 
                                      onFocus={() => setActiveFormulaInput("rootExpression")}
                                      placeholder="x, x + y..."
                                      className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                    />
                                  </div>
                                </div>
                              )}

                              {/* SUMMATION / PRODUCT CONTROLS */}
                              {mathType === "sumatoria" && (
                                <div className="flex flex-col gap-2 animate-fade-in text-[11px]">
                                  <div className="flex items-center gap-4">
                                    <span className="text-zinc-500 font-bold">Tipo:</span>
                                    <label className="flex items-center gap-1 cursor-pointer font-semibold text-zinc-700">
                                      <input 
                                        type="radio" 
                                        checked={sumType === "sum"} 
                                        onChange={() => setSumType("sum")} 
                                        className="accent-emerald-700 h-3.5 w-3.5"
                                      />
                                      Sumatoria (Σ)
                                    </label>
                                    <label className="flex items-center gap-1 cursor-pointer font-semibold text-zinc-700">
                                      <input 
                                        type="radio" 
                                        checked={sumType === "prod"} 
                                        onChange={() => setSumType("prod")} 
                                        className="accent-emerald-700 h-3.5 w-3.5"
                                      />
                                      Productoria (Π)
                                    </label>
                                  </div>

                                  <div className="grid grid-cols-2 gap-2 mt-0.5">
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Límite Inferior:</span>
                                      <input 
                                        type="text" 
                                        value={sumLower} 
                                        onChange={(e) => setSumLower(e.target.value)} 
                                        onFocus={() => setActiveFormulaInput("sumLower")}
                                        placeholder="i=1, n=0, k..."
                                        className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                      />
                                    </div>
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Límite Superior:</span>
                                      <input 
                                        type="text" 
                                        value={sumUpper} 
                                        onChange={(e) => setSumUpper(e.target.value)} 
                                        onFocus={() => setActiveFormulaInput("sumUpper")}
                                        placeholder="n, ∞, 10..."
                                        className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                      />
                                    </div>
                                  </div>

                                  <div className="flex flex-col gap-0.5">
                                    <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Término General:</span>
                                    <input 
                                      type="text" 
                                      value={sumExpression} 
                                      onChange={(e) => setSumExpression(e.target.value)} 
                                      onFocus={() => setActiveFormulaInput("sumExpression")}
                                      placeholder="x_i, i²..."
                                      className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                    />
                                  </div>
                                </div>
                              )}

                              {/* FRACCION CONTROLS */}
                              {mathType === "fraccion" && (
                                <div className="flex flex-col gap-2 animate-fade-in text-[11px]">
                                  <div className="grid grid-cols-2 gap-2">
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Numerador:</span>
                                      <input 
                                        type="text" 
                                        value={fracNumerator} 
                                        onChange={(e) => setFracNumerator(e.target.value)} 
                                        onFocus={() => setActiveFormulaInput("fracNumerator")}
                                        placeholder="Numerador..."
                                        className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                      />
                                    </div>
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Denominador:</span>
                                      <input 
                                        type="text" 
                                        value={fracDenominator} 
                                        onChange={(e) => setFracDenominator(e.target.value)} 
                                        onFocus={() => setActiveFormulaInput("fracDenominator")}
                                        placeholder="Denominador..."
                                        className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                      />
                                    </div>
                                  </div>
                                </div>
                              )}

                              {/* POTENCIA CONTROLS */}
                              {mathType === "potencia" && (
                                <div className="flex flex-col gap-2 animate-fade-in text-[11px]">
                                  <div className="grid grid-cols-2 gap-2">
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Base:</span>
                                      <input 
                                        type="text" 
                                        value={powerBase} 
                                        onChange={(e) => setPowerBase(e.target.value)} 
                                        onFocus={() => setActiveFormulaInput("powerBase")}
                                        placeholder="Base (x)..."
                                        className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                      />
                                    </div>
                                    <div className="flex flex-col gap-0.5">
                                      <span className="text-[9px] text-emerald-800 font-bold uppercase tracking-wider">Exponente (Elevado):</span>
                                      <input 
                                        type="text" 
                                        value={powerExponent} 
                                        onChange={(e) => setPowerExponent(e.target.value)} 
                                        onFocus={() => setActiveFormulaInput("powerExponent")}
                                        placeholder="Exponente (n)..."
                                        className="bg-zinc-50 border border-zinc-250 rounded px-2 py-1 font-semibold text-zinc-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 text-[10px]"
                                      />
                                    </div>
                                  </div>
                                </div>
                              )}
                            </div>

                            {/* Click to insert symbols inside focused field */}
                            <div className="border-t border-emerald-100 pt-2">
                              <span className="text-[8.5px] font-extrabold text-emerald-800 uppercase tracking-wide font-mono block mb-1">
                                {activeFormulaInput 
                                  ? `Insertar en campo activo (${activeFormulaInput === "integralLower" ? "Límite Inferior" : activeFormulaInput === "integralUpper" ? "Límite Superior" : activeFormulaInput === "integralExpression" ? "Expresión" : activeFormulaInput === "rootIndex" ? "Índice" : activeFormulaInput === "rootExpression" ? "Radicando" : activeFormulaInput === "sumLower" ? "Límite Inferior" : activeFormulaInput === "sumUpper" ? "Límite Superior" : activeFormulaInput === "sumExpression" ? "Término" : activeFormulaInput === "fracNumerator" ? "Numerador" : activeFormulaInput === "fracDenominator" ? "Denominador" : activeFormulaInput === "powerBase" ? "Base" : "Exponente"}):`
                                  : "Haz clic en un campo para símbolos rápidos:"}
                              </span>
                              <div className="flex flex-wrap gap-1 max-w-full">
                                {["∞", "π", "√", "θ", "Δ", "α", "β", "γ", "ε", "ω", "σ", "μ", "λ", "φ", "∂", "±", "×", "÷", "≠", "≈", "≤", "≥"].map(sym => (
                                  <button
                                    key={sym}
                                    type="button"
                                    onMouseDown={(e) => e.preventDefault()}
                                    onClick={() => {
                                      if (activeFormulaInput) {
                                        appendSymbolToActiveInput(sym);
                                      } else {
                                        insertCharacterAtCaret(sym);
                                      }
                                    }}
                                    className="w-5.5 h-5.5 flex items-center justify-center hover:bg-emerald-700 hover:text-white border border-emerald-250 rounded text-[10px] font-bold font-mono transition bg-emerald-50 text-emerald-800 cursor-pointer active:scale-90"
                                    title={`Insertar ${sym}`}
                                  >
                                    {sym}
                                  </button>
                                ))}
                              </div>
                            </div>
                          </div>

                          {/* Col 3: Live Preview & Action */}
                          <div className="md:col-span-4 flex flex-col gap-2.5 bg-white border border-emerald-150 rounded-lg p-2.5 items-stretch justify-between">
                            <span className="text-[8px] font-extrabold text-emerald-800 uppercase tracking-wide font-mono">3. Vista Previa & Insertar</span>

                            {/* Whiteboard style live render */}
                            <div className="flex-1 bg-zinc-900 text-white rounded-lg p-3 flex items-center justify-center min-h-[70px] select-none border border-zinc-800 shadow-inner">
                              <div 
                                className="text-center font-serif text-lg selection:bg-transparent"
                                dangerouslySetInnerHTML={{
                                  __html: getMathWrapper(
                                    mathType === "integral" ? getIntegralHtml(integralDefinite, integralLower, integralUpper, integralExpression) :
                                    mathType === "raiz" ? getRootHtml(rootType, rootIndex, rootExpression) :
                                    mathType === "sumatoria" ? getSumProdHtml(sumType, sumLower, sumUpper, sumExpression) :
                                    mathType === "fraccion" ? getFractionHtml(fracNumerator, fracDenominator) :
                                    getPowerHtml(powerBase, powerExponent),
                                    mathSize
                                  )
                                }}
                              />
                            </div>

                            {/* Size selector */}
                            <div className="flex items-center justify-between gap-1.5 mt-1 border-t border-zinc-150 pt-2">
                              <span className="text-[9px] font-bold text-zinc-500 uppercase">Tamaño de fórmula:</span>
                              <div className="flex bg-emerald-50 border border-emerald-250 rounded p-0.5 shrink-0">
                                {["chica", "mediana", "grande"].map(size => (
                                  <button
                                    key={size}
                                    type="button"
                                    onClick={() => setMathSize(size as any)}
                                    className={`px-2 py-0.5 text-[9px] font-bold rounded transition capitalize cursor-pointer
                                      ${mathSize === size 
                                        ? "bg-emerald-700 text-white shadow-xs" 
                                        : "text-zinc-600 hover:bg-emerald-100"}`}
                                  >
                                    {size}
                                  </button>
                                ))}
                              </div>
                            </div>

                            <button
                              type="button"
                              onMouseDown={(e) => e.preventDefault()}
                              onClick={() => {
                                const innerHtml = 
                                  mathType === "integral" ? getIntegralHtml(integralDefinite, integralLower, integralUpper, integralExpression) :
                                  mathType === "raiz" ? getRootHtml(rootType, rootIndex, rootExpression) :
                                  mathType === "sumatoria" ? getSumProdHtml(sumType, sumLower, sumUpper, sumExpression) :
                                  mathType === "fraccion" ? getFractionHtml(fracNumerator, fracDenominator) :
                                  getPowerHtml(powerBase, powerExponent);
                                const html = getMathWrapper(innerHtml, mathSize);
                                insertHtmlAtCaret(html);
                              }}
                              className="w-full bg-emerald-700 hover:bg-emerald-800 active:scale-95 text-white font-bold py-1.5 px-3 rounded-lg text-[10px] transition shadow-md hover:shadow-emerald-700/20 text-center uppercase tracking-wide cursor-pointer"
                            >
                              Insertar en Documento
                            </button>
                          </div>
                        </div>
                      </div>

                    </div>
                  )}

                  {/* SECTION 5: CONFIG PÁG */}
                  {openSections.configPag && (
                    <div className="flex flex-col gap-3.5 text-xs py-2 bg-emerald-50/40 border border-emerald-100 rounded-2xl p-3 shadow-xs animate-fade-in select-none">
                      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-emerald-100 pb-2">
                        <div className="flex items-center gap-2">
                          <span className="px-2 py-0.5 bg-emerald-700 text-white text-[9px] font-extrabold rounded-md uppercase tracking-wider">Gestor de Páginas</span>
                          <span className="text-zinc-600 font-medium">Reordena y edita las páginas de tu documento en miniatura</span>
                        </div>
                        <div className="flex items-center gap-1.5 flex-wrap">
                          <button
                            type="button"
                            onClick={() => {
                              if (selectedPages.length === docPages.length) {
                                setSelectedPages([]);
                              } else {
                                setSelectedPages(docPages.map((_, i) => i));
                              }
                            }}
                            className="px-2 py-1 bg-white hover:bg-zinc-100 border border-zinc-200 text-zinc-700 text-[10px] font-bold rounded-lg transition cursor-pointer"
                          >
                            {selectedPages.length === docPages.length ? "Deseleccionar Todo" : "Seleccionar Todo"}
                          </button>
                          
                          <button
                            type="button"
                            disabled={selectedPages.length === 0}
                            onClick={moveSelectedPagesUp}
                            className={`px-2 py-1 flex items-center gap-1 text-[10px] font-bold rounded-lg border transition ${selectedPages.length > 0 ? "bg-white hover:bg-zinc-100 border-zinc-200 text-zinc-700 cursor-pointer" : "bg-zinc-50 border-zinc-100 text-zinc-300 cursor-not-allowed"}`}
                          >
                            <ChevronLeft className="w-3 h-3" /> Subir ({selectedPages.length})
                          </button>
                          
                          <button
                            type="button"
                            disabled={selectedPages.length === 0}
                            onClick={() => {
                              if (selectedPages.length === 0) return;
                              const sortedIndices = [...selectedPages].sort((a, b) => b - a);
                              if (sortedIndices[0] === docPages.length - 1) return;
                              
                              setDocPages(prev => {
                                const updated = [...prev];
                                for (const idx of sortedIndices) {
                                  const temp = updated[idx];
                                  updated[idx] = updated[idx + 1];
                                  updated[idx + 1] = temp;
                                }
                                return updated;
                              });
                              if (selectedPages.includes(activePageIdx) && activePageIdx < docPages.length - 1) {
                                setActivePageIdx(activePageIdx + 1);
                              }
                              setSelectedPages(prev => prev.map(idx => idx < docPages.length - 1 ? idx + 1 : idx));
                            }}
                            className={`px-2 py-1 flex items-center gap-1 text-[10px] font-bold rounded-lg border transition ${selectedPages.length > 0 ? "bg-white hover:bg-zinc-100 border-zinc-200 text-zinc-700 cursor-pointer" : "bg-zinc-50 border-zinc-100 text-zinc-300 cursor-not-allowed"}`}
                          >
                            Bajar ({selectedPages.length}) <ChevronRight className="w-3 h-3" />
                          </button>

                          <button
                            type="button"
                            disabled={selectedPages.length === 0}
                            onClick={deleteSelectedPages}
                            className={`px-2 py-1 flex items-center gap-1 text-[10px] font-bold rounded-lg border transition ${selectedPages.length > 0 ? "bg-rose-50 hover:bg-rose-100 border-rose-200 text-rose-700 cursor-pointer" : "bg-zinc-50 border-zinc-100 text-zinc-300 cursor-not-allowed"}`}
                          >
                            <Trash2 className="w-3 h-3 text-rose-600" /> Eliminar ({selectedPages.length})
                          </button>
                        </div>
                      </div>

                      <div className="flex gap-4 overflow-x-auto py-3 px-1 scrollbar-thin scrollbar-thumb-emerald-200 scrollbar-track-transparent min-h-[230px]">
                        {docPages.map((pageHtml, idx) => {
                          const isPageSelected = selectedPages.includes(idx);
                          const isPageActive = idx === activePageIdx;
                          
                          return (
                            <div 
                              key={`thumb-page-${idx}`}
                              className="flex flex-col items-center gap-1.5 shrink-0"
                            >
                              <div 
                                onClick={() => {
                                  if (canvasRef.current) {
                                    const currentHtml = canvasRef.current.innerHTML;
                                    setDocPages(prev => {
                                      const updated = [...prev];
                                      updated[activePageIdx] = currentHtml;
                                      return updated;
                                    });
                                  }
                                  setActivePageIdx(idx);
                                }}
                                className={`relative w-[130px] h-[184px] bg-white border rounded-xl shadow-xs transition-all duration-200 overflow-hidden cursor-pointer select-none
                                  ${isPageActive ? "ring-2 ring-emerald-600 border-emerald-500 scale-[1.02] shadow-md" : "border-zinc-200 hover:border-emerald-300"}
                                  ${isPageSelected ? "bg-emerald-50/10 border-emerald-400" : ""}`}
                              >
                                <div 
                                  className="absolute top-0 left-0 origin-top-left p-4 select-none pointer-events-none w-[722px] h-[1022px]"
                                  style={{ 
                                    transform: 'scale(0.18)',
                                    fontFamily: globalFontFamily.startsWith("uploaded-") ? `'${globalFontFamily}', sans-serif` :
                                               globalFontFamily === "font-sans" ? "'Inter', sans-serif" :
                                               globalFontFamily === "font-serif" ? "'Merriweather', Georgia, serif" :
                                               globalFontFamily === "font-mono" ? "'JetBrains Mono', monospace" :
                                               globalFontFamily === "font-display" ? "'Space Grotesk', sans-serif" :
                                               globalFontFamily === "font-elegant" ? "'Playfair Display', serif" :
                                               globalFontFamily === "font-classic" ? "'Cinzel', serif" : "'Inter', sans-serif",
                                    fontSize: globalFontSize,
                                    color: globalTextColor,
                                    backgroundColor: paperBg
                                  }}
                                  dangerouslySetInnerHTML={{ __html: pageHtml || "<p class='italic text-zinc-300'>Página vacía</p>" }}
                                />

                                <div className="absolute inset-x-0 bottom-0 bg-gradient-to-t from-black/60 to-transparent p-1.5 flex items-center justify-between pointer-events-auto">
                                  <input 
                                    type="checkbox"
                                    checked={isPageSelected}
                                    onClick={(e) => {
                                      e.stopPropagation();
                                    }}
                                    onChange={() => {
                                      setSelectedPages(prev => 
                                        prev.includes(idx) ? prev.filter(i => i !== idx) : [...prev, idx]
                                      );
                                    }}
                                    className="w-3.5 h-3.5 rounded border-zinc-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
                                  />
                                  <span className="text-[9px] font-bold text-white tracking-wide">Pág. {idx + 1}</span>
                                </div>
                              </div>

                              <div className="flex items-center gap-1">
                                <button
                                  type="button"
                                  disabled={idx === 0}
                                  onClick={(e) => { e.stopPropagation(); movePageUp(idx); }}
                                  className={`p-1 rounded-md border transition ${idx > 0 ? "bg-white hover:bg-zinc-50 border-zinc-200 text-zinc-700 cursor-pointer" : "bg-zinc-50 border-zinc-100 text-zinc-300 cursor-not-allowed"}`}
                                  title="Subir / Mover Izquierda"
                                >
                                  <ChevronLeft className="w-3 h-3" />
                                </button>
                                
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setDocPages(prev => {
                                      const updated = [...prev];
                                      updated.splice(idx + 1, 0, pageHtml);
                                      return updated;
                                    });
                                  }}
                                  className="p-1 rounded-md bg-white hover:bg-zinc-50 border border-zinc-200 text-zinc-700 transition cursor-pointer"
                                  title="Duplicar página"
                                >
                                  <Copy className="w-3 h-3" />
                                </button>

                                <button
                                  type="button"
                                  disabled={docPages.length <= 1}
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    if (docPages.length <= 1) return;
                                    setDocPages(prev => prev.filter((_, i) => i !== idx));
                                    if (activePageIdx === idx) {
                                      setActivePageIdx(idx > 0 ? idx - 1 : 0);
                                    } else if (activePageIdx > idx) {
                                      setActivePageIdx(activePageIdx - 1);
                                    }
                                  }}
                                  className={`p-1 rounded-md border transition ${docPages.length > 1 ? "bg-rose-50 hover:bg-rose-100 border-rose-200 text-rose-700 cursor-pointer" : "bg-zinc-50 border-zinc-100 text-zinc-300 cursor-not-allowed"}`}
                                  title="Eliminar página"
                                >
                                  <Trash2 className="w-3 h-3" />
                                </button>

                                <button
                                  type="button"
                                  disabled={idx === docPages.length - 1}
                                  onClick={(e) => { e.stopPropagation(); movePageDown(idx); }}
                                  className={`p-1 rounded-md border transition ${idx < docPages.length - 1 ? "bg-white hover:bg-zinc-50 border-zinc-200 text-zinc-700 cursor-pointer" : "bg-zinc-50 border-zinc-100 text-zinc-300 cursor-not-allowed"}`}
                                  title="Bajar / Mover Derecha"
                                >
                                  <ChevronRight className="w-3 h-3" />
                                </button>
                              </div>
                            </div>
                          );
                        })}

                        <div className="flex flex-col items-center justify-center gap-1.5 shrink-0 self-start mt-2">
                          <button
                            type="button"
                            onClick={() => {
                              setDocPages(prev => [...prev, "<p><br></p>"]);
                              setActivePageIdx(docPages.length);
                            }}
                            className="w-[130px] h-[184px] bg-zinc-50 hover:bg-emerald-50/20 border-2 border-dashed border-zinc-300 hover:border-emerald-500 rounded-xl flex flex-col items-center justify-center gap-2 text-zinc-400 hover:text-emerald-700 transition duration-200 cursor-pointer group"
                          >
                            <div className="w-8 h-8 rounded-full bg-zinc-100 group-hover:bg-emerald-100 flex items-center justify-center text-zinc-500 group-hover:text-emerald-700 transition duration-200">
                              <Plus className="w-4 h-4" />
                            </div>
                            <span className="text-[11px] font-bold uppercase tracking-wider">Nueva Pág.</span>
                          </button>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* SECTION 6: AJUSTES */}
                  {openSections.ajustes && (
                    <div className="flex flex-col gap-3 text-xs py-1 animate-fade-in text-zinc-800">
                      <div className="bg-zinc-50/80 border border-zinc-200 rounded-xl p-3 shadow-xs flex flex-col gap-2.5">
                        <div>
                          <span className="text-[10px] font-extrabold text-zinc-500 uppercase tracking-wider font-mono block">Navegación y visualización entre páginas</span>
                          <p className="text-[11px] text-zinc-500 mt-0.5">Selecciona el método de visualización preferido para las páginas de tu documento:</p>
                        </div>
                        
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5 mt-1">
                          {/* Option 1: Scroll Continuo ("scroll") */}
                          <div 
                            onClick={() => setNavigationMode("scroll")}
                            className={`p-3 rounded-xl border-2 transition cursor-pointer select-none flex flex-col justify-between gap-2.5 h-full ${navigationMode === "scroll" ? "bg-emerald-50/20 border-emerald-600" : "bg-white border-zinc-200 hover:border-zinc-300"}`}
                          >
                            <div className="flex items-start gap-2.5">
                              <div className={`p-1.5 rounded-lg border transition ${navigationMode === "scroll" ? "bg-emerald-600 border-emerald-600 text-white" : "bg-zinc-100 border-zinc-200 text-zinc-500"}`}>
                                <MousePointer className="w-4 h-4 animate-pulse" />
                              </div>
                              <div className="flex flex-col">
                                <span className="font-bold text-[11px] text-zinc-850">Vista Continua (Ruedita)</span>
                                <span className="text-[10px] text-zinc-500 mt-0.5 leading-normal">
                                  Como en Microsoft Office. Desplázate verticalmente entre todas las páginas de forma fluida. Se detectará automáticamente la página visible y se animará su selección.
                                </span>
                              </div>
                            </div>
                            <div className="flex items-center gap-1.5 text-[9px] text-emerald-700 font-bold font-mono">
                              {navigationMode === "scroll" && <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>}
                              <span>{navigationMode === "scroll" ? "ACTIVADO" : "HAGA CLIC PARA ACTIVAR"}</span>
                            </div>
                          </div>

                          {/* Option 2: Flechas de Navegación ("pages") */}
                          <div 
                            onClick={() => setNavigationMode("pages")}
                            className={`p-3 rounded-xl border-2 transition cursor-pointer select-none flex flex-col justify-between gap-2.5 h-full ${navigationMode === "pages" ? "bg-emerald-50/20 border-emerald-600" : "bg-white border-zinc-200 hover:border-zinc-300"}`}
                          >
                            <div className="flex items-start gap-2.5">
                              <div className={`p-1.5 rounded-lg border transition ${navigationMode === "pages" ? "bg-emerald-600 border-emerald-600 text-white" : "bg-zinc-100 border-zinc-200 text-zinc-500"}`}>
                                <ChevronRight className="w-4 h-4 animate-bounce" />
                              </div>
                              <div className="flex flex-col">
                                <span className="font-bold text-[11px] text-zinc-850">Páginas Individuales (Flechas)</span>
                                <span className="text-[10px] text-zinc-500 mt-0.5 leading-normal">
                                  Navega de una en una mediante las flechas de pantalla o de teclado. Incluye transiciones de animación fluidas al cambiar de página.
                                </span>
                              </div>
                            </div>
                            <div className="flex items-center gap-1.5 text-[9px] text-emerald-700 font-bold font-mono">
                              {navigationMode === "pages" && <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>}
                              <span>{navigationMode === "pages" ? "ACTIVADO" : "HAGA CLIC PARA ACTIVAR"}</span>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                </div>
              </div>
            )}

            {/* RULER AND MARGIN HELPER RENDER METHODS */}
            {(() => {
              // We define the render methods inside an immediately invoked function expression to avoid scope bleeding
              const renderHorizontalRuler = () => {
                const widthPx = pageDimensions.width * PAGE_SCALE;
                const ticks = [];
                // Generate tick marks for each mm
                for (let i = 0; i <= pageDimensions.width * 10; i++) {
                  const cm = i / 10;
                  const isCm = i % 10 === 0;
                  const isHalfCm = i % 5 === 0;
                  const leftPx = cm * PAGE_SCALE;
                  
                  ticks.push(
                    <div 
                      key={`hr-${i}`}
                      className="absolute bottom-0 bg-zinc-300"
                      style={{
                        left: `${leftPx}px`,
                        width: "1px",
                        height: isCm ? "12px" : isHalfCm ? "8px" : "4px",
                      }}
                    >
                      {isCm && (
                        <span className="absolute left-1/2 -translate-x-1/2 bottom-3.5 text-[7px] text-zinc-500 font-mono font-bold select-none">
                          {cm}
                        </span>
                      )}
                    </div>
                  );
                }

                return (
                  <div 
                    className="relative bg-zinc-50 border-b border-zinc-250 h-6 select-none shadow-xs border-t border-l border-r border-zinc-200 rounded-t-md"
                    style={{ width: `${widthPx}px` }}
                  >
                    {/* Render ticks */}
                    <div className="absolute inset-0 overflow-hidden">
                      {ticks}
                    </div>

                    {/* Drag handle for Margen Izquierdo (Left Margin) */}
                    <div
                      onMouseDown={(e) => handleStartDrag(e, "marginLeft", pageDimensions.width, widthPx)}
                      className="absolute bottom-0 h-full w-3 -ml-1.5 flex flex-col items-center justify-end cursor-col-resize z-30 group"
                      style={{ left: `${marginLeft * PAGE_SCALE}px` }}
                      title="Ajustar Margen Izquierdo"
                    >
                      <div className="w-2 h-3 bg-blue-600 rounded-t-sm shadow-sm group-hover:bg-blue-700 transition"></div>
                      <div className="w-[1px] h-3 bg-blue-400 group-hover:bg-blue-600 transition"></div>
                    </div>

                    {/* Drag handle for Margen Derecho (Right Margin) */}
                    {marginsMode === "4m" && (
                      <div
                        onMouseDown={(e) => handleStartDrag(e, "marginRight", pageDimensions.width, widthPx)}
                        className="absolute bottom-0 h-full w-3 -ml-1.5 flex flex-col items-center justify-end cursor-col-resize z-30 group"
                        style={{ left: `${(pageDimensions.width - marginRight) * PAGE_SCALE}px` }}
                        title="Ajustar Margen Derecho"
                      >
                        <div className="w-2 h-3 bg-blue-600 rounded-t-sm shadow-sm group-hover:bg-blue-700 transition"></div>
                        <div className="w-[1px] h-3 bg-blue-400 group-hover:bg-blue-600 transition"></div>
                      </div>
                    )}

                    {/* Drag handle for Sangría de Párrafo */}
                    <div
                      onMouseDown={(e) => handleStartDrag(e, "indent", pageDimensions.width, widthPx)}
                      className="absolute top-0 w-3 -ml-1.5 flex flex-col items-center justify-start cursor-col-resize z-40 group"
                      style={{ left: `${(marginLeft + textIndent) * PAGE_SCALE}px` }}
                      title="Sangría de Primera Línea"
                    >
                      <div className="w-0 h-0 border-l-[5px] border-l-transparent border-r-[5px] border-r-transparent border-t-[8px] border-t-amber-600 drop-shadow-sm group-hover:border-t-amber-700 transition"></div>
                      <div className="w-[1px] h-3 bg-amber-400 group-hover:bg-amber-600 transition"></div>
                    </div>
                  </div>
                );
              };

              const renderVerticalRuler = () => {
                const heightCm = pageDimensions.height || 29.7;
                const heightPx = heightCm * PAGE_SCALE;
                const ticks = [];
                
                // Generate tick marks for each mm
                for (let i = 0; i <= heightCm * 10; i++) {
                  const cm = i / 10;
                  const isCm = i % 10 === 0;
                  const isHalfCm = i % 5 === 0;
                  const topPx = cm * PAGE_SCALE;
                  
                  ticks.push(
                    <div 
                      key={`vr-${i}`}
                      className="absolute right-0 bg-zinc-300"
                      style={{
                        top: `${topPx}px`,
                        height: "1px",
                        width: isCm ? "12px" : isHalfCm ? "8px" : "4px",
                      }}
                    >
                      {isCm && (
                        <span className="absolute top-1/2 -translate-y-1/2 right-3.5 text-[7px] text-zinc-500 font-mono font-bold select-none">
                          {cm}
                        </span>
                      )}
                    </div>
                  );
                }

                return (
                  <div 
                    className="relative bg-zinc-50 border-r border-zinc-250 select-none shadow-xs border-t border-l border-b border-zinc-200 rounded-l-md"
                    style={{ height: `${heightPx}px`, width: "30px" }}
                  >
                    {/* Render ticks */}
                    <div className="absolute inset-0 overflow-hidden">
                      {ticks}
                    </div>

                    {/* Drag handle for Margen Superior (Top Margin) */}
                    <div
                      onMouseDown={(e) => handleStartDrag(e, "marginTop", heightCm, heightPx)}
                      className="absolute right-0 w-full h-3 -mt-1.5 flex items-center justify-end cursor-row-resize z-30 group"
                      style={{ top: `${marginTop * PAGE_SCALE}px` }}
                      title="Ajustar Margen Superior"
                    >
                      <div className="h-[1px] w-3 bg-blue-400 group-hover:bg-blue-600 transition"></div>
                      <div className="h-2 w-3 bg-blue-600 rounded-l-sm shadow-sm group-hover:bg-blue-700 transition"></div>
                    </div>

                    {/* Drag handle for Margen Inferior (Bottom Margin) */}
                    {pageDimensions.height && marginsMode === "4m" && (
                      <div
                        onMouseDown={(e) => handleStartDrag(e, "marginBottom", heightCm, heightPx)}
                        className="absolute right-0 w-full h-3 -mt-1.5 flex items-center justify-end cursor-row-resize z-30 group"
                        style={{ top: `${(heightCm - marginBottom) * PAGE_SCALE}px` }}
                        title="Ajustar Margen Inferior"
                      >
                        <div className="h-[1px] w-3 bg-blue-400 group-hover:bg-blue-600 transition"></div>
                        <div className="h-2 w-3 bg-blue-600 rounded-l-sm shadow-sm group-hover:bg-blue-700 transition"></div>
                      </div>
                    )}

                    {/* Drag handle for Encabezado Height */}
                    {showHeader && (
                      <div
                        onMouseDown={(e) => handleStartDrag(e, "header", heightCm, heightPx)}
                        className="absolute right-0 w-full h-2.5 -mt-1.25 flex items-center justify-end cursor-row-resize z-40 group"
                        style={{ top: `${headerHeight * PAGE_SCALE}px` }}
                        title="Ajustar Altura de Encabezado"
                      >
                        <div className="h-[1px] w-3 bg-emerald-400 group-hover:bg-emerald-600 transition"></div>
                        <div className="h-1.5 w-3 bg-emerald-600 rounded-l-sm shadow-sm group-hover:bg-emerald-700 transition"></div>
                      </div>
                    )}

                    {/* Drag handle for Pie de Página Height */}
                    {showFooter && pageDimensions.height && (
                      <div
                        onMouseDown={(e) => handleStartDrag(e, "footer", heightCm, heightPx)}
                        className="absolute right-0 w-full h-2.5 -mt-1.25 flex items-center justify-end cursor-row-resize z-40 group"
                        style={{ top: `${(heightCm - footerHeight) * PAGE_SCALE}px` }}
                        title="Ajustar Altura de Pie de Página"
                      >
                        <div className="h-[1px] w-3 bg-emerald-400 group-hover:bg-emerald-600 transition"></div>
                        <div className="h-1.5 w-3 bg-emerald-600 rounded-l-sm shadow-sm group-hover:bg-emerald-700 transition"></div>
                      </div>
                    )}
                  </div>
                );
              };

              const renderGuidelineOverlays = () => {
                const widthPx = pageDimensions.width * PAGE_SCALE;
                const heightCm = pageDimensions.height || 29.7;
                const heightPx = heightCm * PAGE_SCALE;

                return (
                  <div className="absolute inset-0 pointer-events-none z-20">
                    {/* Left margin guideline */}
                    {(isDragging === "marginLeft" || showRulerLines) && (
                      <div 
                        className="absolute top-0 bottom-0 border-l border-dashed border-blue-400 opacity-50"
                        style={{ left: `${marginLeft * PAGE_SCALE}px` }}
                      />
                    )}

                    {/* Right margin guideline */}
                    {marginsMode === "4m" && (isDragging === "marginRight" || showRulerLines) && (
                      <div 
                        className="absolute top-0 bottom-0 border-r border-dashed border-blue-400 opacity-50"
                        style={{ left: `${(pageDimensions.width - marginRight) * PAGE_SCALE}px` }}
                      />
                    )}

                    {/* Text Indentation guideline */}
                    {isDragging === "indent" && (
                      <div 
                        className="absolute top-0 bottom-0 border-l border-dashed border-amber-500 opacity-70"
                        style={{ left: `${(marginLeft + textIndent) * PAGE_SCALE}px` }}
                      />
                    )}

                    {/* Top margin guideline */}
                    {(isDragging === "marginTop" || showRulerLines) && (
                      <div 
                        className="absolute left-0 right-0 border-t border-dashed border-blue-400 opacity-50"
                        style={{ top: `${marginTop * PAGE_SCALE}px` }}
                      />
                    )}

                    {/* Bottom margin guideline */}
                    {pageDimensions.height && marginsMode === "4m" && (isDragging === "marginBottom" || showRulerLines) && (
                      <div 
                        className="absolute left-0 right-0 border-b border-dashed border-blue-400 opacity-50"
                        style={{ top: `${(heightCm - marginBottom) * PAGE_SCALE}px` }}
                      />
                    )}

                    {/* Header guideline */}
                    {showHeader && (isDragging === "header" || showRulerLines) && (
                      <div 
                        className="absolute left-0 right-0 border-t border-dashed border-emerald-400 opacity-50"
                        style={{ top: `${headerHeight * PAGE_SCALE}px` }}
                      />
                    )}

                    {/* Footer guideline */}
                    {showFooter && pageDimensions.height && (isDragging === "footer" || showRulerLines) && (
                      <div 
                        className="absolute left-0 right-0 border-t border-dashed border-emerald-400 opacity-50"
                        style={{ top: `${(heightCm - footerHeight) * PAGE_SCALE}px` }}
                      />
                    )}
                  </div>
                );
              };

              return (
                <div className="w-full flex flex-col items-center gap-2">
                  <div className="text-[10px] text-zinc-400 font-bold uppercase tracking-widest font-mono flex items-center gap-1 pl-1 selection:bg-none w-full justify-center">
                    <Eye className="w-3.5 h-3.5 text-zinc-800 animate-pulse" />
                    <span>Haz clic directo en cualquier parte de la hoja para editar o redactar</span>
                  </div>

                  {/* Dragging indicator toast */}
                  {isDragging && (
                    <div className="px-3.5 py-1.5 bg-zinc-900 text-white text-[10px] font-bold rounded-full shadow-lg border border-zinc-800 flex items-center gap-2 animate-bounce">
                      <span className="w-2 h-2 rounded-full bg-blue-500 animate-ping"></span>
                      <span>Ajustando {
                        isDragging === "marginLeft" ? `Margen Izquierdo a ${marginLeft} cm` :
                        isDragging === "marginRight" ? `Margen Derecho a ${marginRight} cm` :
                        isDragging === "marginTop" ? `Margen Superior a ${marginTop} cm` :
                        isDragging === "marginBottom" ? `Margen Inferior a ${marginBottom} cm` :
                        isDragging === "indent" ? `Sangría de Párrafo a ${textIndent} cm` :
                        isDragging === "header" ? `Altura de Encabezado a ${headerHeight} cm` :
                        `Altura de Pie de Página a ${footerHeight} cm`
                      }</span>
                    </div>
                  )}

                  {/* Complete Workspace with centered Rulers and dynamic paper sheets */}
                  <div 
                    onWheel={handleWheelAddPage}
                    onScroll={handleWorksheetScroll}
                    className="flex-1 w-full bg-zinc-100/60 rounded-3xl p-6 border border-zinc-200/80 shadow-inner flex flex-col items-center justify-start min-h-[600px] overflow-auto relative"
                  >
                    <div className="flex flex-col items-center relative" style={{ width: "fit-content" }}>
                      
                      {/* Horizontal Ruler */}
                      <div className="flex select-none" style={{ paddingLeft: "46px" }}>
                        {renderHorizontalRuler()}
                      </div>

                      <div className="flex items-start mt-5 relative">
                        {/* Vertical Ruler on the Left side */}
                        <div className="sticky top-0 shrink-0" style={{ width: "30px", marginRight: "16px" }}>
                          {renderVerticalRuler()}
                        </div>

                        {/* Interactive Paper container */}
                        {navigationMode === "scroll" ? (
                          <div className="flex flex-col items-center">
                            {docPages.map((pageHtml, i) => {
                              const isActive = i === activePageIdx;
                              return (
                                <div 
                                  key={`paper-sheet-scroll-${i}`}
                                  id={`paper-sheet-container-${i}`}
                                  onClick={() => setActivePageIdx(i)}
                                  className={`relative bg-white border shadow-md transition-all duration-300 origin-center ${isActive ? 'ring-2 ring-emerald-600 border-emerald-500 scale-[1.01] shadow-xl z-10' : 'border-zinc-300/80 opacity-90 hover:opacity-100'}`}
                                  style={{ 
                                    width: `${pageDimensions.width * PAGE_SCALE}px`,
                                    height: pageDimensions.height ? `${pageDimensions.height * PAGE_SCALE}px` : "auto",
                                    minHeight: pageDimensions.height ? `${pageDimensions.height * PAGE_SCALE}px` : "960px",
                                    marginBottom: "32px"
                                  }}
                                >
                                  {isActive && renderGuidelineOverlays()}

                                  {/* Editable paper content sheet with padding margins */}
                                  <div
                                    style={{
                                      backgroundColor: paperBg,
                                      color: globalTextColor,
                                      fontFamily: getFontFamilyString(globalFontFamily),
                                      fontSize: globalFontSize,
                                      paddingTop: i === 0 && !coverHasMargins ? "0cm" : `${marginTop}cm`,
                                      paddingBottom: i === 0 && !coverHasMargins ? "0cm" : `${marginBottom}cm`,
                                      paddingLeft: i === 0 && !coverHasMargins ? "0cm" : `${marginLeft}cm`,
                                      paddingRight: i === 0 && !coverHasMargins ? "0cm" : `${marginRight}cm`,
                                      width: "100%",
                                      height: "100%",
                                      minHeight: pageDimensions.height ? `${pageDimensions.height * PAGE_SCALE}px` : "960px",
                                    }}
                                    className="relative text-left flex flex-col justify-between"
                                  >
                                    {/* Header (Absolute positioned in the margin area) */}
                                    {showHeader && !(i === 0 && !coverHasMargins) && (
                                      <div 
                                        style={{ 
                                          position: "absolute",
                                          top: `${headerHeight}cm`,
                                          left: `${marginLeft}cm`,
                                          right: `${marginRight}cm`,
                                          color: headerColor, 
                                          fontSize: headerSize,
                                          textAlign: headerAlign,
                                          fontWeight: headerBold ? "bold" : "normal",
                                          fontStyle: headerItalic ? "italic" : "normal"
                                        }}
                                        className="border-b border-zinc-200 pb-1.5 uppercase tracking-widest font-mono select-none"
                                      >
                                        {docHeader || <span className="text-zinc-350 italic">Encabezado de página</span>}
                                      </div>
                                    )}

                                    {/* Main Document Text Area */}
                                    <div 
                                      id={isActive ? "rich-doc-canvas" : `rich-doc-canvas-inactive-${i}`}
                                      ref={isActive ? canvasRef : null}
                                      contentEditable
                                      suppressContentEditableWarning
                                      onMouseUp={saveSelection}
                                      onKeyUp={saveSelection}
                                      onKeyDown={handleKeyDownAddPage}
                                      onInput={(e) => {
                                        const newHtml = e.currentTarget.innerHTML;
                                        setDocPages(prev => {
                                          const updated = [...prev];
                                          updated[i] = newHtml;
                                          return updated;
                                        });
                                        if (isActive) {
                                          setEditedHtml(newHtml);
                                        }
                                      }}
                                      onBlur={(e) => {
                                        const newHtml = e.currentTarget.innerHTML;
                                        setDocPages(prev => {
                                          const updated = [...prev];
                                          updated[i] = newHtml;
                                          return updated;
                                        });
                                        if (isActive) {
                                          setEditedHtml(newHtml);
                                        }
                                      }}
                                      dangerouslySetInnerHTML={{ __html: pageHtml || "<p><br></p>" }}
                                      className="outline-none min-h-[760px] h-full prose prose-sm max-w-none text-inherit focus:outline-none flex-1
                                        prose-headings:font-bold prose-headings:text-inherit prose-headings:mt-4 prose-headings:mb-2
                                        prose-h1:text-xl prose-h2:text-lg prose-h3:text-base prose-p:mb-3 prose-p:mt-1
                                        prose-ul:list-disc prose-ul:pl-5 prose-ol:list-decimal prose-ol:pl-5 prose-li:mb-1"
                                    />

                                    {/* Footer (Absolute positioned in the margin area) */}
                                    {(showFooter || showPageNumbers) && !(i === 0 && !coverHasMargins) && (
                                      <div 
                                        style={{ 
                                          position: "absolute",
                                          bottom: `${footerHeight}cm`,
                                          left: `${marginLeft}cm`,
                                          right: `${marginRight}cm`,
                                          color: footerColor, 
                                          fontSize: footerSize,
                                          fontWeight: footerBold ? "bold" : "normal",
                                          fontStyle: footerItalic ? "italic" : "normal"
                                        }}
                                        className="border-t border-zinc-200 pt-2 uppercase tracking-widest font-mono select-none"
                                      >
                                        <div className="w-full text-[10px]">
                                          {showFooter && (
                                            <div style={{ textAlign: footerAlign }} className="text-zinc-500 w-full mb-1">
                                              {docFooter}
                                            </div>
                                          )}
                                          {showPageNumbers && (
                                            <div style={{ textAlign: pageNumberAlign }} className="text-zinc-600 font-bold w-full">
                                              {getPageNumberText(i)}
                                            </div>
                                          )}
                                        </div>
                                      </div>
                                    )}

                                  </div>
                                </div>
                              );
                            })}

                            {/* Wide white button with black text/border at the end of scroll list */}
                            {pageSize !== "Infinita" && (
                              <button
                                type="button"
                                onClick={handleAddNewPage}
                                className="py-4 bg-white hover:bg-zinc-50 border-2 border-dashed border-zinc-300 hover:border-zinc-400 text-zinc-950 hover:text-black font-extrabold rounded-2xl text-[11px] transition duration-200 shadow-md flex items-center justify-center gap-2 cursor-pointer uppercase tracking-widest z-10"
                                style={{ width: `${pageDimensions.width * PAGE_SCALE}px` }}
                              >
                                <Plus className="w-4 h-4 text-zinc-850 stroke-[3]" /> Agregar otra página
                              </button>
                            )}
                          </div>
                        ) : (
                          <div className="flex flex-col items-center gap-6">
                            {/* Animatable single sheet */}
                            <AnimatePresence mode="wait">
                              <motion.div
                                key={`paper-sheet-pages-${activePageIdx}`}
                                initial={{ opacity: 0, x: 40 }}
                                animate={{ opacity: 1, x: 0 }}
                                exit={{ opacity: 0, x: -40 }}
                                transition={{ duration: 0.25, ease: "easeInOut" }}
                                className="relative shadow-2xl border border-zinc-300 bg-white"
                                style={{ 
                                  width: `${pageDimensions.width * PAGE_SCALE}px`,
                                  height: pageDimensions.height ? `${pageDimensions.height * PAGE_SCALE}px` : "auto",
                                  minHeight: pageDimensions.height ? `${pageDimensions.height * PAGE_SCALE}px` : "960px",
                                }}
                              >
                                {renderGuidelineOverlays()}

                                {/* Editable paper content sheet with padding margins */}
                                <div
                                  style={{
                                    backgroundColor: paperBg,
                                    color: globalTextColor,
                                    fontFamily: getFontFamilyString(globalFontFamily),
                                    fontSize: globalFontSize,
                                    paddingTop: activePageIdx === 0 && !coverHasMargins ? "0cm" : `${marginTop}cm`,
                                    paddingBottom: activePageIdx === 0 && !coverHasMargins ? "0cm" : `${marginBottom}cm`,
                                    paddingLeft: activePageIdx === 0 && !coverHasMargins ? "0cm" : `${marginLeft}cm`,
                                    paddingRight: activePageIdx === 0 && !coverHasMargins ? "0cm" : `${marginRight}cm`,
                                    width: "100%",
                                    height: "100%",
                                    minHeight: pageDimensions.height ? `${pageDimensions.height * PAGE_SCALE}px` : "960px",
                                  }}
                                  className="relative text-left flex flex-col justify-between"
                                >
                                  {/* Header */}
                                  {showHeader && !(activePageIdx === 0 && !coverHasMargins) && (
                                    <div 
                                      style={{ 
                                        position: "absolute",
                                        top: `${headerHeight}cm`,
                                        left: `${marginLeft}cm`,
                                        right: `${marginRight}cm`,
                                        color: headerColor, 
                                        fontSize: headerSize,
                                        textAlign: headerAlign,
                                        fontWeight: headerBold ? "bold" : "normal",
                                        fontStyle: headerItalic ? "italic" : "normal"
                                      }}
                                      className="border-b border-zinc-200 pb-1.5 uppercase tracking-widest font-mono select-none"
                                    >
                                      {docHeader || <span className="text-zinc-350 italic">Encabezado de página</span>}
                                    </div>
                                  )}

                                  {/* Main Document Text Area */}
                                  <div 
                                    id="rich-doc-canvas"
                                    ref={canvasRef}
                                    contentEditable
                                    suppressContentEditableWarning
                                    onMouseUp={saveSelection}
                                    onKeyUp={saveSelection}
                                    onKeyDown={handleKeyDownAddPage}
                                    onInput={(e) => {
                                      const newHtml = e.currentTarget.innerHTML;
                                      setDocPages(prev => {
                                        const updated = [...prev];
                                        updated[activePageIdx] = newHtml;
                                        return updated;
                                      });
                                      setEditedHtml(newHtml);
                                    }}
                                    onBlur={(e) => {
                                      const newHtml = e.currentTarget.innerHTML;
                                      setDocPages(prev => {
                                        const updated = [...prev];
                                        updated[activePageIdx] = newHtml;
                                        return updated;
                                      });
                                      setEditedHtml(newHtml);
                                    }}
                                    dangerouslySetInnerHTML={{ __html: docPages[activePageIdx] || "<p><br></p>" }}
                                    className="outline-none min-h-[760px] h-full prose prose-sm max-w-none text-inherit focus:outline-none flex-1
                                      prose-headings:font-bold prose-headings:text-inherit prose-headings:mt-4 prose-headings:mb-2
                                      prose-h1:text-xl prose-h2:text-lg prose-h3:text-base prose-p:mb-3 prose-p:mt-1
                                      prose-ul:list-disc prose-ul:pl-5 prose-ol:list-decimal prose-ol:pl-5 prose-li:mb-1"
                                  />

                                  {/* Footer */}
                                  {(showFooter || showPageNumbers) && !(activePageIdx === 0 && !coverHasMargins) && (
                                    <div 
                                      style={{ 
                                        position: "absolute",
                                        bottom: `${footerHeight}cm`,
                                        left: `${marginLeft}cm`,
                                        right: `${marginRight}cm`,
                                        color: footerColor, 
                                        fontSize: footerSize,
                                        fontWeight: footerBold ? "bold" : "normal",
                                        fontStyle: footerItalic ? "italic" : "normal"
                                      }}
                                      className="border-t border-zinc-200 pt-2 uppercase tracking-widest font-mono select-none"
                                    >
                                      <div className="w-full text-[10px]">
                                        {showFooter && (
                                          <div style={{ textAlign: footerAlign }} className="text-zinc-500 w-full mb-1">
                                            {docFooter}
                                          </div>
                                        )}
                                        {showPageNumbers && (
                                          <div style={{ textAlign: pageNumberAlign }} className="text-zinc-600 font-bold w-full">
                                            {getPageNumberText(activePageIdx)}
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  )}
                                </div>
                              </motion.div>
                            </AnimatePresence>

                            {/* Unified, gorgeous compact navigation pill below the sheet in Pages mode */}
                            <div className="flex items-center gap-3 bg-zinc-950 text-white px-5 py-2.5 rounded-2xl shadow-2xl border border-zinc-800/80 animate-fade-in z-20 select-none">
                              <button
                                type="button"
                                disabled={activePageIdx === 0}
                                onClick={goToPrevPage}
                                className="p-1 hover:bg-zinc-850 rounded-lg text-zinc-300 hover:text-white disabled:opacity-35 disabled:hover:bg-transparent disabled:hover:text-zinc-300 transition cursor-pointer flex items-center gap-1 font-extrabold text-[11px] uppercase tracking-wider"
                                title="Página Anterior"
                              >
                                <ChevronLeft className="w-4 h-4 stroke-[3]" /> Anterior
                              </button>
                              <div className="w-px h-5 bg-zinc-850" />
                              <span className="text-[11px] font-mono font-black text-zinc-400 uppercase tracking-widest px-2">
                                Pág. <span className="text-emerald-400 font-extrabold">{activePageIdx + 1}</span> de <span className="text-zinc-500 font-extrabold">{docPages.length}</span>
                              </span>
                              <div className="w-px h-5 bg-zinc-850" />
                              <button
                                type="button"
                                onClick={goToNextPage}
                                className="p-1 hover:bg-zinc-850 rounded-lg text-zinc-300 hover:text-white transition cursor-pointer flex items-center gap-1 font-extrabold text-[11px] uppercase tracking-wider"
                                title={activePageIdx === docPages.length - 1 ? "Agregar otra página" : "Página Siguiente"}
                              >
                                {activePageIdx === docPages.length - 1 ? (
                                  <>
                                    <Plus className="w-4 h-4 text-emerald-400 stroke-[3]" /> Nueva
                                  </>
                                ) : (
                                  <>
                                    Siguiente <ChevronRight className="w-4 h-4 stroke-[3]" />
                                  </>
                                )}
                              </button>
                            </div>
                          </div>
                        )}

                      </div>
                    </div>
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* MODAL DE SELECCIÓN Y CONFIGURACIÓN DE PORTADAS */}
        {showCoverSelectorModal && (
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[9999] p-4 animate-fade-in">
            <div className="bg-white rounded-3xl w-full max-w-3xl max-h-[90vh] overflow-y-auto shadow-2xl border border-zinc-200 p-6 md:p-8 flex flex-col gap-6">
              
              {/* Modal Header */}
              <div className="flex items-center justify-between border-b border-zinc-150 pb-4">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 bg-emerald-50 rounded-xl text-emerald-600">
                    <Sparkles className="w-5 h-5 animate-pulse" />
                  </div>
                  <div>
                    <h3 className="text-base font-extrabold text-zinc-900">Configurar e Insertar Portada</h3>
                    <p className="text-[11px] text-zinc-500 font-medium font-sans">Selecciona un estilo de hoja completa o importa tus propios códigos CSS de diseño</p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setShowCoverSelectorModal(false)}
                  className="p-1.5 hover:bg-zinc-100 rounded-lg text-zinc-400 hover:text-zinc-600 transition cursor-pointer"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Cover Templates Grid (Previews/Miniatures) */}
              <div>
                <span className="text-[10px] font-extrabold text-zinc-400 uppercase tracking-widest block mb-3 font-mono">1. Miniaturas de Diseños Disponibles</span>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                  
                  {/* Modern Option Card */}
                  <button
                    type="button"
                    onClick={() => {
                      handleInsertCover("modern");
                      setShowCoverSelectorModal(false);
                    }}
                    className="group text-left border border-zinc-200 hover:border-emerald-500 hover:shadow-md rounded-2xl p-3 bg-zinc-50 hover:bg-white transition flex flex-col gap-2 cursor-pointer active:scale-97"
                  >
                    <div className="w-full h-24 rounded-lg bg-gradient-to-br from-zinc-100 to-zinc-200 border border-zinc-200 overflow-hidden relative flex flex-col justify-between p-2">
                      <div className="w-1.5 h-full bg-emerald-600 absolute left-0 top-0"></div>
                      <span className="text-[7px] font-bold text-emerald-700 uppercase tracking-wider pl-2 block">REPORTE</span>
                      <div className="pl-2 flex flex-col gap-0.5">
                        <span className="text-[8px] font-extrabold text-zinc-800 leading-none">TRANSFORMACIÓN</span>
                        <span className="text-[6px] text-zinc-400 leading-none">Un análisis profundo...</span>
                      </div>
                    </div>
                    <div className="flex flex-col">
                      <span className="text-xs font-bold text-zinc-850 group-hover:text-emerald-700">Moderna</span>
                      <span className="text-[10px] text-zinc-400">Izquierda verde</span>
                    </div>
                  </button>

                  {/* Editorial Option Card */}
                  <button
                    type="button"
                    onClick={() => {
                      handleInsertCover("editorial");
                      setShowCoverSelectorModal(false);
                    }}
                    className="group text-left border border-zinc-200 hover:border-emerald-500 hover:shadow-md rounded-2xl p-3 bg-zinc-50 hover:bg-white transition flex flex-col gap-2 cursor-pointer active:scale-97"
                  >
                    <div className="w-full h-24 rounded-lg bg-zinc-50 border border-zinc-200 overflow-hidden relative flex flex-col justify-between p-2 text-center">
                      <span className="text-[7px] font-bold text-zinc-400 uppercase tracking-wider block">MEMORIA ANUAL</span>
                      <span className="text-[9px] font-serif text-zinc-900 block font-bold leading-none my-1">ESTRATEGIA</span>
                      <span className="text-[6px] text-zinc-400 block leading-none">Sinergias Corporativas</span>
                    </div>
                    <div className="flex flex-col">
                      <span className="text-xs font-bold text-zinc-850 group-hover:text-emerald-700">Editorial</span>
                      <span className="text-[10px] text-zinc-400">Minimalista serif</span>
                    </div>
                  </button>

                  {/* Tech Option Card */}
                  <button
                    type="button"
                    onClick={() => {
                      handleInsertCover("tech");
                      setShowCoverSelectorModal(false);
                    }}
                    className="group text-left border border-zinc-200 hover:border-emerald-500 hover:shadow-md rounded-2xl p-3 bg-zinc-50 hover:bg-white transition flex flex-col gap-2 cursor-pointer active:scale-97"
                  >
                    <div className="w-full h-24 rounded-lg bg-zinc-900 border border-zinc-950 overflow-hidden relative flex flex-col justify-between p-2">
                      <span className="text-[6px] font-mono text-emerald-400">STATUS: LIVE</span>
                      <span className="text-[8px] font-mono text-zinc-200 font-bold leading-none">SISTEMAS AUTÓNOMOS</span>
                      <span className="text-[5px] font-mono text-zinc-500">Architecture v2</span>
                    </div>
                    <div className="flex flex-col">
                      <span className="text-xs font-bold text-zinc-850 group-hover:text-emerald-700">Tecnológica</span>
                      <span className="text-[10px] text-zinc-400">Monospace dark</span>
                    </div>
                  </button>

                  {/* Academic Option Card */}
                  <button
                    type="button"
                    onClick={() => {
                      handleInsertCover("academic");
                      setShowCoverSelectorModal(false);
                    }}
                    className="group text-left border border-zinc-200 hover:border-emerald-500 hover:shadow-md rounded-2xl p-3 bg-zinc-50 hover:bg-white transition flex flex-col gap-2 cursor-pointer active:scale-97"
                  >
                    <div className="w-full h-24 rounded-lg bg-white border border-zinc-300 overflow-hidden relative flex flex-col justify-between p-2 text-center border-t-4 border-t-indigo-900">
                      <span className="text-[6px] font-serif text-zinc-600 block">UTN REGIONAL</span>
                      <span className="text-[8px] font-serif text-indigo-950 block font-bold leading-none">ANÁLISIS COMP.</span>
                      <span className="text-[5px] text-zinc-400 block">Juan Ignacio Negroni</span>
                    </div>
                    <div className="flex flex-col">
                      <span className="text-xs font-bold text-zinc-850 group-hover:text-emerald-700">Académica</span>
                      <span className="text-[10px] text-zinc-400">Times New UTN</span>
                    </div>
                  </button>

                </div>
              </div>

              {/* Style Importer Section */}
              <div className="bg-zinc-50 border border-zinc-150 rounded-2xl p-4 md:p-5 flex flex-col gap-4">
                <div>
                  <h4 className="text-xs font-extrabold text-zinc-850 flex items-center gap-1.5 uppercase tracking-wider font-mono">
                    <Sliders className="w-3.5 h-3.5 text-blue-600" />
                    2. Importador de Estilos Portada (Pega CSS)
                  </h4>
                  <p className="text-[11px] text-zinc-500 mt-1">Importa colores, gradientes, bordes o fuentes personalizando el fondo con formato CSS válido en el campo abajo:</p>
                </div>

                <div className="flex flex-col gap-2">
                  <textarea
                    rows={2}
                    value={customCoverStyleString}
                    onChange={(e) => setCustomCoverStyleString(e.target.value)}
                    placeholder="Ej: background: radial-gradient(circle, #fbcfe8 0%, #f472b6 100%); color: #000; border-top: 15px solid #000;"
                    className="w-full bg-white border border-zinc-250 rounded-xl p-3 text-[11px] font-mono focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 text-zinc-800"
                  />
                  
                  {/* Pre-designed presets list to import instantly */}
                  <div className="flex flex-wrap items-center gap-1.5 mt-1">
                    <span className="text-[9px] font-bold text-zinc-400">Presets Rápidos:</span>
                    {[
                      { label: "Gradiente Sunset", style: "background: linear-gradient(135deg, #ff7e5f 0%, #feb47b 100%); color: white; border-top: 10px solid #333;" },
                      { label: "Brutalismo Amarillo", style: "background: #facc15; border: 6px solid #000; color: #000; font-family: Impact, sans-serif;" },
                      { label: "Nordic Minimal", style: "background: #f4f4f5; border: 1px solid #e4e4e7; color: #18181b;" },
                      { label: "Profundo Premium", style: "background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%); border-left: 12px solid #0284c7; color: #f8fafc;" }
                    ].map(preset => (
                      <button
                        key={preset.label}
                        type="button"
                        onClick={() => setCustomCoverStyleString(preset.style)}
                        className="px-2 py-1 bg-white hover:bg-zinc-150 border border-zinc-200 hover:border-zinc-300 rounded text-[9px] font-semibold text-zinc-700 transition active:scale-95 cursor-pointer"
                      >
                        {preset.label}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Import insertion trigger */}
                <div className="flex justify-end">
                  <button
                    type="button"
                    onClick={() => {
                      handleInsertCover("custom");
                      setShowCoverSelectorModal(false);
                    }}
                    className="px-4 py-2 bg-blue-650 hover:bg-blue-700 active:scale-95 text-white font-bold rounded-xl text-xs transition cursor-pointer flex items-center gap-1.5 shadow-md shadow-blue-500/10"
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    Generar Portada Importada
                  </button>
                </div>
              </div>

              {/* Cover Margin Option Controls */}
              <div className="flex items-center justify-between bg-zinc-50 border border-zinc-150 rounded-2xl p-4">
                <div className="flex flex-col gap-0.5">
                  <span className="text-xs font-bold text-zinc-800">¿Márgenes de página en la portada?</span>
                  <span className="text-[10px] text-zinc-500 font-sans">Determina si la portada respeta los márgenes de página configurados o si es de hoja completa (0cm de márgenes).</span>
                </div>
                <label className="relative inline-flex items-center cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={coverHasMargins}
                    onChange={(e) => setCoverHasMargins(e.target.checked)}
                    className="sr-only peer"
                  />
                  <div className="w-10 h-6 bg-zinc-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-zinc-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-emerald-600"></div>
                </label>
              </div>

              {/* Info text */}
              <p className="text-[10px] text-zinc-400 italic text-center">La portada se insertará al principio del documento (como página 1) y será completamente editable haciendo clic directo.</p>

            </div>
          </div>
        )}
      </div>

      {/* SOLID, FIXED BOTTOM OFFICE STATUS BAR */}
      {(isWord || isPDF) && (
        <div className="w-full bg-zinc-900 text-zinc-300 border-t border-zinc-800 px-6 py-2.5 flex flex-wrap items-center justify-between gap-4 text-xs font-medium z-50 shrink-0 select-none">
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-blue-500 animate-pulse"></span>
              <span className="text-zinc-400">Páginas Totales:</span>
              <span className="font-bold text-white font-mono">{docPages.length}</span>
            </div>
            <div className="flex items-center gap-2 border-l border-zinc-800 pl-6">
              <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
              <span className="text-zinc-400">Página Actual:</span>
              <div className="flex items-center gap-1.5 bg-zinc-800 border border-zinc-750 rounded-lg px-2 py-0.5 ml-1 select-none">
                <button
                  type="button"
                  disabled={activePageIdx === 0}
                  onClick={goToPrevPage}
                  className="p-0.5 hover:bg-zinc-700 active:scale-90 rounded text-zinc-300 hover:text-white disabled:opacity-30 disabled:hover:bg-transparent disabled:hover:text-zinc-300 disabled:cursor-not-allowed transition cursor-pointer"
                  title="Página Anterior"
                >
                  <ChevronLeft className="w-3.5 h-3.5" />
                </button>
                <span className="font-bold text-white font-mono px-1 text-[11px]">{activePageIdx + 1} de {docPages.length}</span>
                <button
                  type="button"
                  onClick={goToNextPage}
                  className="p-0.5 hover:bg-zinc-700 active:scale-90 rounded text-zinc-300 hover:text-white transition cursor-pointer"
                  title={activePageIdx === docPages.length - 1 ? "Agregar Nueva Página" : "Página Siguiente"}
                >
                  {activePageIdx === docPages.length - 1 ? (
                    <Plus className="w-3.5 h-3.5 text-emerald-450 stroke-[2.5]" />
                  ) : (
                    <ChevronRight className="w-3.5 h-3.5" />
                  )}
                </button>
              </div>
            </div>
          </div>
          
          <div className="flex items-center gap-6 text-[11px] font-mono">
            <div className="bg-zinc-850 border border-zinc-800 rounded-lg px-3 py-1 flex items-center gap-1.5 text-zinc-300">
              <span className="text-zinc-500">Palabras (Pág / Total):</span>
              <span className="font-bold text-emerald-400">{stats.activeWords}</span>
              <span className="text-zinc-600">/</span>
              <span className="text-zinc-300">{stats.totalWords}</span>
            </div>

            <div className="bg-zinc-850 border border-zinc-800 rounded-lg px-3 py-1 flex items-center gap-1.5 text-zinc-300">
              <span className="text-zinc-500">Caracteres (Pág / Total):</span>
              <span className="font-bold text-emerald-400">{stats.activeChars}</span>
              <span className="text-zinc-600">/</span>
              <span className="text-zinc-300">{stats.totalChars}</span>
            </div>
            
            <div className="bg-zinc-850 border border-zinc-850 rounded-lg px-2.5 py-1 text-zinc-400 font-sans font-bold flex items-center gap-1.5 uppercase text-[9px] tracking-wider">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-ping"></span>
              <span>Modo: {navigationMode === "scroll" ? "Continuo (Rueda)" : "Individual (Flechas)"}</span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
