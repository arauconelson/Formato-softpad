/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from "react";
import { useDropzone } from "react-dropzone";
import { 
  FileText, 
  Upload, 
  Download, 
  Table as TableIcon, 
  Plus, 
  X, 
  Loader2, 
  CheckCircle2, 
  FileSpreadsheet,
  AlertCircle,
  Trash2,
  Files
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import ExcelJS from "exceljs";

import { cn } from "./lib/utils";
import { DocRow } from "./types";
import { extractTableFromImage, FileData } from "./services/geminiService";

export default function App() {
  const [rows, setRows] = useState<DocRow[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [loadingMessage, setLoadingMessage] = useState("");
  const [complementaryFiles, setComplementaryFiles] = useState<FileData[]>([]);
  const [mainFile, setMainFile] = useState<FileData | null>(null);

  const loadingMessages = [
    "Analizando estructura del documento...",
    "Identificando columnas de la plantilla...",
    "Contrastando datos con archivos complementarios...",
    "Generando códigos ISO de país...",
    "Normalizando estado de productos...",
    "Validando correspondencia de campos...",
    "Casi listo, el modelo está refinando los detalles...",
  ];

  const handleProcess = async () => {
    if (!mainFile) return;

    setIsLoading(true);
    setError(null);
    
    // Cycle messages
    let msgIdx = 0;
    const interval = setInterval(() => {
      msgIdx = (msgIdx + 1) % loadingMessages.length;
      setLoadingMessage(loadingMessages[msgIdx]);
    }, 2500);
    setLoadingMessage(loadingMessages[0]);

    try {
      const extractedRows = await extractTableFromImage(mainFile, complementaryFiles);
      setRows(prev => [...prev, ...extractedRows]);
      setMainFile(null); // Clear main file after processing
    } catch (err: any) {
      setError("Error al procesar el documento. Asegúrate de que la imagen sea clara.");
      console.error(err);
    } finally {
      setIsLoading(false);
      clearInterval(interval);
    }
  };

  const onDrop = useCallback(async (acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
      setMainFile({ base64: reader.result as string, mimeType: file.type, name: file.name });
    };
    reader.readAsDataURL(file);
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'image/*': ['.jpeg', '.jpg', '.png', '.webp'],
      'application/pdf': ['.pdf']
    },
    multiple: false
  });

  const onDropComplementary = useCallback(async (acceptedFiles: File[]) => {
    acceptedFiles.forEach(file => {
      const reader = new FileReader();
      reader.onload = () => {
        const base64 = reader.result as string;
        setComplementaryFiles(prev => [...prev, { base64, mimeType: file.type, name: file.name }]);
      };
      reader.readAsDataURL(file);
    });
  }, []);

  const { getRootProps: getCompRootProps, getInputProps: getCompInputProps, isDragActive: isCompDragActive } = useDropzone({
    onDrop: onDropComplementary,
    accept: {
      'image/*': ['.jpeg', '.jpg', '.png', '.webp'],
      'application/pdf': ['.pdf']
    },
    multiple: true
  });

  const removeCompFile = (idx: number) => {
    setComplementaryFiles(prev => prev.filter((_, i) => i !== idx));
  };

  const updateRow = (id: string, field: keyof DocRow, value: string) => {
    setRows(prev => prev.map(row => row.id === id ? { ...row, [field]: value } : row));
  };

  const deleteRow = (id: string) => {
    setRows(prev => prev.filter(row => row.id !== id));
  };

  const clearAllRows = () => {
    setRows([]);
  };

  const addRow = () => {
    const newRow: DocRow = {
      id: crypto.randomUUID(),
      proveedor: "",
      factura: "",
      fechaFactura: "",
      item: "",
      partida: "",
      marca: "",
      modelo: "",
      cantidadUnidadComercial: "",
      tipoUnidadComercial: "U",
      cantidadBultos: "",
      claseDeBultos: "BUL",
      precioItemUSD: "",
      paisOrigen: "",
      estado: "10",
      pesoNeto: "",
      numeroParte: "",
      descripcion: "Nuevo ítem",
      caracteristica1: "",
      caracteristica2: "",
      caracteristica3: "",
      caracteristica4: "",
      confidenceMap: {}
    };
    setRows([...rows, newRow]);
  };

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Documentos');

    const headers = [
      "PROVEEDOR", 
      "FACTURA", 
      "FECHA DE FACTURA", 
      "ÍTEM", 
      "PARTIDA", 
      "MARCA", 
      "MODELO", 
      "CANTIDAD UNIDAD COMERCIAL", 
      "TIPO DE UNIDAD COMERCIAL", 
      "CANTIDAD BULTOS", 
      "CLASE DE BULTOS", 
      "PRECIO ITEM US$", 
      "PAIS ORIGEN", 
      "ESTADO", 
      "PESO NETO", 
      "NUMERO PARTE", 
      "DESCRIPCIÓN", 
      "CARACTERISTICA 1", 
      "CARACTERISTICA 2", 
      "CARACTERISTICA 3", 
      "CARACTERISTICA 4"
    ];

    const headerRow = worksheet.addRow(headers);
    
    // Header style
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E293B' } // slate-800
      };
      cell.font = {
        color: { argb: 'FFFFFFFF' },
        bold: true,
        size: 10
      };
      cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    const colKeys: (keyof DocRow)[] = [
      'proveedor', 'factura', 'fechaFactura', 'item', 'partida', 'marca', 'modelo',
      'cantidadUnidadComercial', 'tipoUnidadComercial', 'cantidadBultos', 'claseDeBultos',
      'precioItemUSD', 'paisOrigen', 'estado', 'pesoNeto', 'numeroParte', 
      'descripcion', 'caracteristica1', 'caracteristica2', 'caracteristica3', 'caracteristica4'
    ];

    rows.forEach(row => {
      const dataRow = worksheet.addRow(colKeys.map(key => row[key] || ""));
      
      dataRow.eachCell((cell, colNumber) => {
        const key = colKeys[colNumber - 1];
        const value = row[key] || "";
        const confidence = row.confidenceMap?.[key] || 'high';

        let bgColor = 'FFFFFFFF'; // White

        if (value === "") {
          bgColor = 'FFFFE4E1'; // light coral / red-50
        } else if (confidence === 'medium') {
          bgColor = 'FFFFF9C4'; // amber-50
        } else if (confidence === 'low') {
          bgColor = 'FFFFE4E1'; // red-50
        } else {
          bgColor = 'FFF0FDF4'; // green-50
        }

        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: bgColor }
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };
        cell.font = { size: 9 };
      });
    });

    // Auto fit column width (rough estimate)
    worksheet.columns.forEach(column => {
      column.width = 20;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Importacion_DocuFlow_${new Date().getTime()}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const columns: { key: keyof DocRow; label: string }[] = [
    { key: "proveedor", label: "Proveedor" },
    { key: "factura", label: "Factura" },
    { key: "fechaFactura", label: "Fecha Factura" },
    { key: "item", label: "Ítem" },
    { key: "partida", label: "Partida" },
    { key: "marca", label: "Marca" },
    { key: "modelo", label: "Modelo" },
    { key: "cantidadUnidadComercial", label: "Cant. Unid. Com." },
    { key: "tipoUnidadComercial", label: "Tipo Unid. Com." },
    { key: "cantidadBultos", label: "Cant. Bultos" },
    { key: "claseDeBultos", label: "Clase Bultos" },
    { key: "precioItemUSD", label: "Precio USD" },
    { key: "paisOrigen", label: "País Origen" },
    { key: "estado", label: "Estado" },
    { key: "pesoNeto", label: "Peso Neto" },
    { key: "numeroParte", label: "Nº Parte" },
    { key: "descripcion", label: "Descripción" },
    { key: "caracteristica1", label: "Característica 1" },
    { key: "caracteristica2", label: "Característica 2" },
    { key: "caracteristica3", label: "Característica 3" },
    { key: "caracteristica4", label: "Característica 4" },
  ];

  return (
    <div className="min-h-screen flex flex-col bg-slate-50">
      {/* Top Navigation Bar */}
      <nav className="h-16 bg-white border-b border-slate-200 px-6 flex items-center justify-between shadow-sm z-10 sticky top-0">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center text-white font-bold text-lg">A</div>
          <span className="text-xl font-semibold tracking-tight text-slate-800">
            DocuFlow <span className="text-blue-600">Pro</span>
          </span>
        </div>
        
        <div className="flex items-center gap-6">
          <div className="flex gap-4 text-sm font-medium text-slate-600">
            <span className="text-blue-600 border-b-2 border-blue-600 pb-5 pt-5 cursor-pointer">Panel de Control</span>
          </div>
          
          <div className="flex items-center gap-3 border-l border-slate-200 pl-6">
            <div className="flex items-center gap-4 mr-4">
              <div className="flex items-center gap-1.5 grayscale opacity-70">
                <div className="w-3 h-3 bg-green-500 rounded-full border border-white"></div>
                <span className="text-[10px] font-bold">100%</span>
              </div>
              <div className="flex items-center gap-1.5 grayscale opacity-70">
                <div className="w-3 h-3 bg-amber-400 rounded-full border border-white"></div>
                <span className="text-[10px] font-bold">DUDOSO</span>
              </div>
              <div className="flex items-center gap-1.5 grayscale opacity-70">
                <div className="w-3 h-3 bg-red-400 rounded-full border border-white"></div>
                <span className="text-[10px] font-bold">VACÍO</span>
              </div>
            </div>
            {rows.length > 0 && (
              <div className="flex items-center gap-2">
                <button 
                  onClick={exportToExcel}
                  className="px-4 py-2 text-[11px] font-bold bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-all shadow-md shadow-blue-100 flex items-center gap-2 uppercase tracking-wide active:scale-95"
                  id="export-excel"
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  DESCARGAR EXCEL (.XLSX)
                </button>
              </div>
            )}
          </div>
        </div>
      </nav>

      <main className="flex-1 flex flex-col overflow-hidden p-6 max-w-[1600px] mx-auto w-full gap-6">
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6 shrink-0">
          {/* Central Workspace: Drag & Drop Area */}
          <section 
            {...getRootProps()} 
            className={cn(
              "lg:col-span-3 relative group border-2 border-dashed rounded-xl p-12 transition-all cursor-pointer overflow-hidden flex flex-col items-center justify-center bg-white shadow-sm",
              isDragActive ? "border-blue-600 bg-blue-50/30" : "border-slate-200 bg-white",
              (isLoading || !!mainFile) && "pointer-events-none opacity-50",
              mainFile && "border-blue-600 bg-blue-50/10 border-solid"
            )}
          >
            <input {...getInputProps()} />
            
            <div className="relative z-10 flex flex-col items-center">
              <div className={cn(
                "p-4 rounded-xl mb-4 transition-all duration-300",
                (isDragActive || mainFile) ? "bg-blue-100 text-blue-600 scale-110" : "bg-slate-50 text-slate-400 group-hover:scale-105"
              )}>
                {isLoading ? (
                  <Loader2 className="w-8 h-8 animate-spin" />
                ) : mainFile ? (
                  <CheckCircle2 className="w-8 h-8 text-blue-600" />
                ) : (
                  <Upload className="w-8 h-8" />
                )}
              </div>
              <h2 className={cn("text-lg font-bold tracking-tight", mainFile ? "text-blue-700" : "text-slate-800")}>
                {mainFile ? mainFile.name : (isDragActive ? "Suelta para cargar" : "Factura o Póliza Principal")}
              </h2>
              <p className="text-slate-500 text-xs mt-1 max-w-sm text-center">
                {mainFile ? "Documento listo para procesar" : "Arrastra aquí el documento que contiene la tabla de datos."}
              </p>
              {mainFile && (
                <button 
                  onClick={(e) => { e.stopPropagation(); setMainFile(null); }}
                  className="mt-4 px-3 py-1 bg-white border border-slate-200 text-slate-500 rounded-md text-[10px] font-bold hover:bg-red-50 hover:text-red-500 transition-all pointer-events-auto"
                >
                  CAMBIAR ARCHIVO
                </button>
              )}
            </div>
            
            {isLoading && (
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="absolute inset-0 bg-white/90 backdrop-blur-[2px] z-20 flex flex-col items-center justify-center"
              >
                <div className="w-64 h-2 bg-slate-100 rounded-full overflow-hidden mb-6">
                  <motion.div 
                    initial={{ x: "-100%" }}
                    animate={{ x: "0%" }}
                    transition={{ duration: 2.5, repeat: Infinity, ease: "easeInOut" }}
                    className="w-full h-full bg-blue-600"
                  />
                </div>
                <p className="font-mono text-[10px] uppercase tracking-[0.2em] text-slate-800 font-bold animate-pulse text-center px-4">
                  {loadingMessage}
                </p>
              </motion.div>
            )}
          </section>

          {/* Complementary Files Section */}
          <section className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm flex flex-col gap-4">
            <header className="flex items-center gap-2 border-b border-slate-100 pb-3">
              <Files className="w-4 h-4 text-blue-600" />
              <h3 className="text-xs font-bold text-slate-800 uppercase tracking-wider">Archivos de Contexto</h3>
            </header>

            <div 
              {...getCompRootProps()} 
              className={cn(
                "border-2 border-dashed rounded-lg p-4 transition-all cursor-pointer flex flex-col items-center justify-center gap-2",
                isCompDragActive ? "border-blue-500 bg-blue-50" : "border-slate-100 hover:border-blue-200 bg-slate-50"
              )}
            >
              <input {...getCompInputProps()} />
              <Upload className="w-4 h-4 text-slate-400" />
              <span className="text-[10px] text-center font-medium text-slate-600">Subir catálogos o fichas técnicas</span>
            </div>

            <div className="flex-1 overflow-auto max-h-[120px] custom-scrollbar">
              <div className="flex flex-col gap-2">
                <AnimatePresence>
                  {complementaryFiles.map((file, idx) => (
                    <motion.div 
                      key={idx}
                      initial={{ opacity: 0, y: 5 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, x: 20 }}
                      className="flex items-center justify-between p-2 bg-slate-50 border border-slate-100 rounded-md group"
                    >
                      <div className="flex items-center gap-2 truncate">
                        <FileText className="w-3 h-3 text-slate-400 shrink-0" />
                        <span className="text-[10px] font-medium text-slate-700 truncate">{file.name || `Archivo ${idx + 1}`}</span>
                      </div>
                      <button 
                        onClick={() => removeCompFile(idx)}
                        className="p-1 hover:bg-red-50 text-slate-300 hover:text-red-500 transition-colors"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </motion.div>
                  ))}
                  {complementaryFiles.length === 0 && (
                    <div className="py-8 text-center border border-dashed border-slate-100 rounded-lg">
                      <p className="text-[9px] text-slate-400 px-4">Agrega archivos para ayudar a la IA con las partidas arancelarias</p>
                    </div>
                  )}
                </AnimatePresence>
              </div>
            </div>
          </section>
        </div>

        {mainFile && !isLoading && (
            <motion.div 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="flex flex-col items-center gap-3 bg-blue-600/5 border border-blue-200 rounded-xl p-6"
            >
                <div className="flex items-center gap-3">
                    <div className="flex -space-x-2">
                        <div className="w-8 h-8 rounded-full bg-blue-600 flex items-center justify-center text-white ring-2 ring-white">
                            <FileText className="w-4 h-4" />
                        </div>
                        {complementaryFiles.map((_, i) => (
                            <div key={i} className="w-8 h-8 rounded-full bg-slate-200 flex items-center justify-center text-slate-500 ring-2 ring-white text-[10px] font-bold">
                                +1
                            </div>
                        ))}
                    </div>
                    <p className="text-sm font-semibold text-slate-700">
                        {complementaryFiles.length > 0 
                            ? `Listo para procesar factura con ${complementaryFiles.length} archivos de apoyo`
                            : "Listo para procesar la factura principal"
                        }
                    </p>
                </div>
                <button 
                    onClick={handleProcess}
                    className="px-8 py-3 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold text-sm shadow-xl shadow-blue-200 transition-all active:scale-95 flex items-center gap-2"
                >
                    <Upload className="w-4 h-4" />
                    PROCESAR DOCUMENTOS E INICIAR EXTRACCIÓN
                </button>
            </motion.div>
        )}

        {error && (
          <div className="bg-red-50 border border-red-200 p-4 rounded-lg flex items-center gap-3 text-red-700 shadow-sm border-l-4 border-l-red-500">
            <AlertCircle className="w-5 h-5 flex-shrink-0" />
            <p className="text-sm font-semibold">{error}</p>
          </div>
        )}

        {/* Results Panel */}
        <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-sm flex flex-col flex-1 min-h-[500px]">
          <header className="h-14 bg-slate-50 border-b border-slate-200 flex items-center px-6 justify-between shrink-0">
             <div className="flex items-center gap-4">
                <div className="flex items-center gap-2">
                  <TableIcon className="w-4 h-4 text-slate-500" />
                  <span className="text-xs font-bold uppercase tracking-wide text-slate-500">Datos Extraídos</span>
                </div>
                <div className="h-4 w-[1px] bg-slate-200"></div>
                <div className="text-xs font-semibold text-slate-800">
                  <span className="text-blue-600">{rows.length}</span> ítems detectados
                </div>
             </div>
             
             <div className="flex items-center gap-4">
                <button 
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      clearAllRows();
                    }}
                    disabled={rows.length === 0}
                    className="text-red-500 text-[10px] font-bold hover:underline flex items-center gap-1.5 disabled:opacity-30 disabled:no-underline transition-opacity"
                >
                    <Trash2 className="w-3 h-3" />
                    BORRAR TODO
                </button>
                <div className="h-4 w-[1px] bg-slate-200"></div>
                <button 
                  onClick={addRow}
                  className="text-blue-600 text-[10px] font-bold hover:underline flex items-center gap-1.5"
                >
                  <Plus className="w-3 h-3" />
                  AÑADIR FILA MANUAL
                </button>
             </div>
          </header>

          <div className="flex-1 overflow-auto">
            <table className="w-full text-left border-collapse min-w-[2200px]">
              <thead className="sticky top-0 z-20">
                <tr className="bg-white border-b border-slate-200">
                  <th className="p-4 technical-label w-12 text-center bg-slate-50/50">#</th>
                  {columns.map(col => (
                    <th key={col.key} className="p-4 technical-label min-w-[180px] bg-slate-50/50 border-r border-slate-100 last:border-r-0">
                      {col.label}
                    </th>
                  ))}
                  <th className="p-4 w-12 bg-slate-50/50"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                <AnimatePresence>
                  {rows.length === 0 ? (
                    <tr>
                      <td colSpan={columns.length + 2} className="p-32 text-center">
                        <div className="flex flex-col items-center text-slate-400">
                          <TableIcon className="w-12 h-12 mb-4 opacity-20" />
                          <p className="text-sm italic">Esperando procesamiento de documentos...</p>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    rows.map((row, idx) => (
                      <motion.tr 
                        key={row.id}
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        exit={{ opacity: 0, x: -10 }}
                        className="group hover:bg-slate-50 transition-colors"
                      >
                        <td className="p-4 text-[10px] font-mono text-slate-400 text-center font-bold border-r border-slate-50">{idx + 1}</td>
                        {columns.map(col => {
                          const value = row[col.key] || "";
                          const confidence = row.confidenceMap?.[col.key] || 'high';
                          
                          let bgClass = "bg-white";
                          let borderClass = "border-slate-200";
                          
                          if (value === "") {
                            bgClass = "bg-red-50/50";
                            borderClass = "border-red-200";
                          } else if (confidence === 'medium') {
                            bgClass = "bg-amber-50/50";
                            borderClass = "border-amber-200";
                          } else if (confidence === 'low') {
                            bgClass = "bg-red-50/50";
                            borderClass = "border-red-200";
                          } else {
                            bgClass = "bg-green-50/30";
                            borderClass = "border-green-200/50";
                          }

                          return (
                            <td key={col.key} className="p-2 border-r border-slate-50 last:border-r-0">
                              <input 
                                type="text" 
                                value={value} 
                                onChange={(e) => updateRow(row.id, col.key, e.target.value)}
                                className={cn(
                                  "w-full border text-xs rounded p-2 outline-none focus:ring-1 focus:ring-blue-100 transition-all font-medium text-slate-700",
                                  bgClass,
                                  borderClass
                                )}
                              />
                            </td>
                          );
                        })}
                        <td className="p-4 text-center">
                          <button 
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              deleteRow(row.id);
                            }}
                            className="p-2 rounded-md bg-slate-50 hover:bg-red-50 text-slate-300 hover:text-red-500 transition-all"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </motion.tr>
                    ))
                  )}
                </AnimatePresence>
              </tbody>
            </table>
          </div>
        </div>
      </main>

      {/* Bottom Status Bar */}
      <footer className="h-8 bg-white border-t border-slate-200 px-6 flex items-center justify-between text-[10px] font-bold uppercase tracking-wider text-slate-400">
        <div className="flex gap-6">
          <span className="flex items-center gap-1.5">
            <span className="w-2 h-2 bg-green-500 rounded-full"></span> 
            OCR Engine: Connected
          </span>
          <span className="flex items-center gap-1.5">
            <span className="w-2 h-2 bg-blue-500 rounded-full"></span> 
            Gemini Flash Preview: Ready
          </span>
        </div>
        <div className="flex gap-4 italic font-normal text-slate-300">
          <span>Design: Professional Polish</span>
          <span>© 2026 DocuFlow Pro</span>
        </div>
      </footer>
    </div>
  );
}

