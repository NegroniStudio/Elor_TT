import express from "express";
import path from "path";
import fs from "fs";
import { createServer as createViteServer } from "vite";
import multer from "multer";
import * as pdfParseInstance from "pdf-parse";
const pdfParse = (pdfParseInstance as any).default || pdfParseInstance;
import mammoth from "mammoth";
import * as XLSX from "xlsx";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

// Initialize Express
const app = express();
const PORT = 3000;

// Increase request size limit for large base64 strings if needed
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ limit: "50mb", extended: true }));

// Setup Multer for memory storage file uploads (up to 25MB)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 }, // 25MB limit
});

// Simple legacy .doc text extraction helper
function extractTextFromOldDoc(buffer: Buffer): string {
  // .doc is binary OLE. Text is stored inside contiguous streams.
  // We can scan and extract printable UTF-8/Ascii string sequences.
  const str = buffer.toString("utf8");
  
  // Regex to extract readable blocks: letters, digits, spaces, common punctuation, Spanish accents/questions
  const matchedBlocks = str.match(/[\x20-\x7E\xA0-\xFF\s\u00C0-\u00FF\u0152\u0153\u0178]{15,}/g);
  
  if (!matchedBlocks || matchedBlocks.length === 0) {
    return "No se pudo extraer texto estructurado directo de este archivo .doc binario heredado de Word. Por favor, intente con un formato .docx.";
  }
  
  // Let's filter out blocks with dense binary fragments or long spaces
  return matchedBlocks
    .map(block => block.trim())
    .map(block => block.replace(/\s+/g, " "))
    .filter(block => block.length > 20 && !(/[\x7F-\x9F]/.test(block)))
    .slice(0, 150) // Limit length to avoid extreme trash blocks
    .join("\n\n");
}

// Document Parser Controller
async function parseDocument(file: Express.Multer.File) {
  const extension = path.extname(file.originalname).toLowerCase();
  let extractedText = "";
  let visualContent: any = null;

  switch (extension) {
    case ".pdf": {
      const pdfData = await pdfParse(file.buffer);
      extractedText = pdfData.text || "";
      visualContent = {
        pages: pdfData.numpages,
        version: pdfData.version,
        info: pdfData.info,
        excerpt: extractedText.trim().substring(0, 10000), // Excerpt for preview
      };
      break;
    }

    case ".docx": {
      // Extract raw text for Gemini AI
      const rawTextResult = await mammoth.extractRawText({ buffer: file.buffer });
      extractedText = rawTextResult.value;
      
      // Convert to HTML for gorgeous visual representation in frontend
      const htmlResult = await mammoth.convertToHtml({ buffer: file.buffer });
      visualContent = {
        html: htmlResult.value,
        messagesCount: htmlResult.messages.length,
      };
      break;
    }

    case ".doc": {
      extractedText = extractTextFromOldDoc(file.buffer);
      visualContent = {
        html: `<div class="font-sans whitespace-pre-wrap leading-relaxed text-gray-800 p-4 bg-amber-50/40 rounded-lg border border-amber-100/50">
          <p class="text-xs font-semibold text-amber-800 mb-2 font-mono">⚠️ FORMATO BINARIO DE WORD (.doc) DETECTADO</p>
          ${extractedText.replace(/\n\n/g, "</p><p class='mt-4'>")}
        </div>`,
        excerpt: extractedText,
      };
      break;
    }

    case ".xlsx":
    case ".xls": {
      const workbook = XLSX.read(file.buffer, { type: "buffer" });
      const sheets: any[] = [];
      
      workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        // Read spreadsheet as a grid representation (array of arrays)
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

    default:
      throw new Error(`Extensión de archivo no soportada: ${extension}`);
  }

  // Double guard text size to avoid overloading LLM limits
  let clipped = false;
  const wordLimit = 80000; // Large token window safety
  const rawWords = extractedText.split(/\s+/);
  if (rawWords.length > wordLimit) {
    extractedText = rawWords.slice(0, wordLimit).join(" ");
    clipped = true;
  }

  return {
    extractedText,
    visualContent,
    clipped,
    filename: file.originalname,
    size: file.size,
    extension,
  };
}

// --- API ROUTES ---

// Endpoint to upload and parse documents
app.post("/api/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No se proporcionó ningún archivo." });
    }

    const docData = await parseDocument(req.file);

    res.json({
      success: true,
      filename: docData.filename,
      size: docData.size,
      extension: docData.extension,
      visualContent: docData.visualContent,
      extractedText: docData.extractedText,
      clipped: docData.clipped,
    });
  } catch (err: any) {
    console.error("Error processing upload:", err);
    res.status(500).json({ error: err.message || "Error al procesar el documento." });
  }
});

// Setup server bundle endpoints and frontend files route
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    // Development Mode leveraging Vite as a middleware
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Production Mode serving fully compiled frontend bundle assets from /dist
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  // Launch service
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Document Parse Dev Server is running at http://localhost:${PORT}`);
  });
}

startServer();
